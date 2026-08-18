import streamlit as st
import pandas as pd
import io
import re
import requests

from pptx_utils import extract_text_units, apply_translations, extract_reference_texts
from ai_utils import (
    classify_text_units,
    translate_presentation_texts,
    generate_copy_options,
    chat_modify_presentation,
    chat_refine_copy,
)

st.set_page_config(page_title="PPT 번역 시스템", layout="wide")

# ── Session state initialisation ─────────────────────────────────────────────

def _init():
    defaults = {
        "stage": "upload",
        "file_bytes": None,
        "file_name": "translated.pptx",
        # brief
        "brand_name_ko": "",
        "brand_name_en": "",
        "key_phrases": [],       # list of {"한국어": str, "영어": str}
        "ref_pptx_texts": [],    # English lines from a reference PPTX
        "glossary": "",          # free-text proper nouns to keep as-is
        "font_name": "",         # family name extracted from uploaded font file
        "text_units": [],
        "classified_units": [],
        "classification_done": False,
        "presentation_units": [],
        "copy_units": [],
        "presentation_translations": {},
        "translations_loaded": False,
        "chat_history_2a": [],
        "copy_options": {},
        "copy_selections": {},
        "copy_options_loaded": False,
        "current_copy_idx": 0,
        "chat_history_2b": [],  # list of (instruction, refined_text) per current line
        "output_bytes": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init()

# ── Stage progress indicator ──────────────────────────────────────────────────

STAGES = ["upload", "classify", "review_2a", "review_2b", "download"]
STAGE_LABELS = {
    "upload": "1. 업로드",
    "classify": "2. 분류",
    "review_2a": "3A. 발표용 감수",
    "review_2b": "3B. 카피 선택",
    "download": "4. 다운로드",
}

def _show_progress():
    current = st.session_state.stage
    cols = st.columns(len(STAGES))
    for col, s in zip(cols, STAGES):
        label = STAGE_LABELS[s]
        if s == current:
            col.markdown(f"**:blue[{label}]**")
        elif STAGES.index(s) < STAGES.index(current):
            col.markdown(f"~~{label}~~")
        else:
            col.markdown(f"<span style='color:grey'>{label}</span>", unsafe_allow_html=True)
    st.divider()

_show_progress()


def _build_context() -> str:
    """Assemble the full brief/context string to inject into every AI prompt."""
    parts = []

    brand_ko = st.session_state.brand_name_ko.strip()
    brand_en = st.session_state.brand_name_en.strip()
    if brand_en or brand_ko:
        label = f"'{brand_en}' (EN) / '{brand_ko}' (KO)" if brand_en and brand_ko else brand_en or brand_ko
        parts.append(
            f"Brand name: {label}. Always use '{brand_en or brand_ko}' in English translations."
        )

    kp = [p for p in st.session_state.key_phrases if p.get("한국어", "").strip() and p.get("영어", "").strip()]
    if kp:
        lines = "\n".join(f"  '{p['한국어']}' → '{p['영어']}'" for p in kp)
        parts.append(f"Preferred term translations (use these exact English expressions):\n{lines}")

    glossary = st.session_state.glossary.strip()
    if glossary:
        parts.append(f"Do NOT translate these proper nouns — keep as-is: {glossary}")

    ref = st.session_state.ref_pptx_texts
    if ref:
        sample = "\n".join(f"  - {t}" for t in ref[:30])
        parts.append(
            "Style & terminology reference (from a previous approved English translation — "
            f"follow this vocabulary and register):\n{sample}"
        )

    return "\n\n".join(parts)


# ── Stage: upload ─────────────────────────────────────────────────────────────

def _gdrive_file_id(url: str) -> str | None:
    """Extract Google Drive file ID from various share URL formats."""
    patterns = [
        r"/file/d/([a-zA-Z0-9_-]+)",
        r"[?&]id=([a-zA-Z0-9_-]+)",
        r"/presentation/d/([a-zA-Z0-9_-]+)",
    ]
    for p in patterns:
        m = re.search(p, url)
        if m:
            return m.group(1)
    return None


def _gdrive_is_slides(url: str) -> bool:
    return bool(re.search(r"/presentation/d/", url))


def _download_gdrive(file_id: str, is_slides: bool = False) -> bytes:
    """Download a file from Google Drive, trying multiple strategies."""
    session = requests.Session()
    errors = []

    # Strategy 1: Google Slides export (for Slides files or as fallback)
    if is_slides:
        url = f"https://docs.google.com/presentation/d/{file_id}/export/pptx"
        try:
            resp = session.get(url, timeout=300)
            if resp.status_code == 200 and len(resp.content) > 1000:
                return resp.content
            errors.append(f"Slides export: HTTP {resp.status_code}")
        except Exception as e:
            errors.append(f"Slides export: {e}")

    # Strategy 2: new usercontent endpoint with confirm=t
    url = f"https://drive.usercontent.google.com/download?id={file_id}&export=download&authuser=0&confirm=t"
    try:
        resp = session.get(url, stream=True, timeout=300)
        if resp.status_code == 200 and len(resp.content) > 1000:
            return resp.content
        errors.append(f"usercontent: HTTP {resp.status_code}")
    except Exception as e:
        errors.append(f"usercontent: {e}")

    # Strategy 3: try Slides export even if URL didn't look like Slides
    if not is_slides:
        url = f"https://docs.google.com/presentation/d/{file_id}/export/pptx"
        try:
            resp = session.get(url, timeout=300)
            if resp.status_code == 200 and len(resp.content) > 1000:
                return resp.content
            errors.append(f"Slides fallback: HTTP {resp.status_code}")
        except Exception as e:
            errors.append(f"Slides fallback: {e}")

    # Strategy 4: legacy uc endpoint with cookie confirmation
    url = f"https://drive.google.com/uc?export=download&id={file_id}"
    try:
        resp = session.get(url, stream=True, timeout=300)
        for key, value in resp.cookies.items():
            if key.startswith("download_warning"):
                url = f"https://drive.google.com/uc?export=download&id={file_id}&confirm={value}"
                resp = session.get(url, stream=True, timeout=300)
                break
        if resp.status_code == 200 and len(resp.content) > 1000:
            return resp.content
        errors.append(f"legacy uc: HTTP {resp.status_code}")
    except Exception as e:
        errors.append(f"legacy uc: {e}")

    raise RuntimeError(f"모든 다운로드 방법 실패: {'; '.join(errors)}")


if st.session_state.stage == "upload":
    st.header("PPT 파일 업로드 및 캠페인 브리프")

    # ── 1. PPTX 파일 ──────────────────────────────────────────────────────────
    st.subheader("한국어 PPTX 파일")
    tab_local, tab_drive = st.tabs(["📁 직접 업로드", "☁️ Google Drive 링크"])

    with tab_local:
        uploaded = st.file_uploader("PPTX 파일 선택", type=["pptx"], label_visibility="collapsed")

    with tab_drive:
        st.caption("파일을 Google Drive에서 공유(링크가 있는 모든 사용자 → 뷰어)한 뒤 링크를 붙여넣으세요.")
        drive_url = st.text_input("Google Drive 공유 링크", placeholder="https://drive.google.com/file/d/.../view?usp=sharing", label_visibility="collapsed")
        drive_bytes = None
        drive_name = None
        if drive_url.strip():
            fid = _gdrive_file_id(drive_url.strip())
            if fid:
                st.caption(f"파일 ID: `{fid}`")
            else:
                st.warning("올바른 Google Drive 링크가 아닙니다.")

    # Resolve whichever source the user provided
    def _resolve_file():
        if uploaded is not None:
            return uploaded.read(), uploaded.name
        if drive_url.strip():
            fid = _gdrive_file_id(drive_url.strip())
            if not fid:
                return None, None
            with st.spinner("Google Drive에서 파일 다운로드 중..."):
                try:
                    data = _download_gdrive(fid, is_slides=_gdrive_is_slides(drive_url.strip()))
                    name = f"gdrive_{fid[:8]}.pptx"
                    return data, name
                except Exception as e:
                    st.error(f"다운로드 실패: {e}")
                    return None, None
        return None, None

    file_ready = (uploaded is not None) or bool(drive_url.strip() and _gdrive_file_id(drive_url.strip()))

    st.divider()

    # ── 2. 브랜드명 ───────────────────────────────────────────────────────────
    st.subheader("브랜드명")
    col_ko, col_en = st.columns(2)
    with col_ko:
        brand_ko = st.text_input("한국어", placeholder="예: 삼성전자", value=st.session_state.brand_name_ko)
    with col_en:
        brand_en = st.text_input("영어", placeholder="예: Samsung Electronics", value=st.session_state.brand_name_en)

    st.divider()

    # ── 3. 주요 용어 한↔영 매핑 ──────────────────────────────────────────────
    st.subheader("주요 용어 매핑 (한국어 → 영어)")
    st.caption("자주 쓰는 표현의 선호 번역을 지정합니다. 행 추가 버튼으로 항목을 추가하세요.")

    init_kp = st.session_state.key_phrases if st.session_state.key_phrases else [{"한국어": "", "영어": ""}]
    kp_df = pd.DataFrame(init_kp)
    edited_kp = st.data_editor(
        kp_df,
        column_config={
            "한국어": st.column_config.TextColumn("한국어 표현", width="large"),
            "영어": st.column_config.TextColumn("영어 번역 (선호)", width="large"),
        },
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        key="kp_editor",
    )

    st.divider()

    # ── 4. 이전 번역본 참고 (선택) ────────────────────────────────────────────
    st.subheader("이전 번역본 참고 (선택)")
    st.caption("기존 영문 PPT를 올리면 용어·문체를 참고해 일관성을 유지합니다.")
    ref_uploaded = st.file_uploader("영문 참고 PPTX (선택)", type=["pptx"], key="ref_uploader")

    st.divider()

    # ── 5. 영어 폰트 (선택) ───────────────────────────────────────────────────
    st.subheader("영어 폰트 (선택)")
    st.caption("번역된 텍스트에 적용할 TTF/OTF 폰트 파일을 업로드하세요. (해당 폰트가 PPT를 열 PC에 설치되어 있어야 합니다)")
    font_uploaded = st.file_uploader("폰트 파일 (선택, TTF/OTF)", type=["ttf", "otf"], key="font_uploader")
    if font_uploaded is not None:
        st.caption(f"업로드된 파일: {font_uploaded.name}")

    st.divider()

    # ── 6. 고유명사 (번역 유지) ───────────────────────────────────────────────
    st.subheader("고유명사 / 번역하지 않을 단어")
    st.caption("영어로도 그대로 쓸 브랜드명, 인명, 제품명 등을 쉼표로 구분해 입력하세요.")
    glossary = st.text_input(
        "고유명사",
        placeholder="예: ChatGPT, POSCO, K-Beauty",
        value=st.session_state.glossary,
        label_visibility="collapsed",
    )

    st.divider()

    if st.button("분석 시작", type="primary", disabled=not file_ready):
        file_bytes, raw_name = _resolve_file()
        if file_bytes is None:
            st.error("파일을 불러오지 못했습니다.")
            st.stop()
        with st.spinner("텍스트 파싱 중..."):
            text_units = extract_text_units(file_bytes)

        if not text_units:
            st.error("번역 가능한 텍스트를 찾지 못했습니다. 파일을 확인해 주세요.")
            st.stop()

        # Save brief to session state
        st.session_state.brand_name_ko = brand_ko
        st.session_state.brand_name_en = brand_en
        st.session_state.key_phrases = edited_kp.to_dict("records")
        st.session_state.glossary = glossary

        if ref_uploaded is not None:
            with st.spinner("참고 번역본 분석 중..."):
                st.session_state.ref_pptx_texts = extract_reference_texts(ref_uploaded.read())
        else:
            st.session_state.ref_pptx_texts = []

        if font_uploaded is not None:
            try:
                from fontTools import ttLib
                tt = ttLib.TTFont(io.BytesIO(font_uploaded.read()))
                extracted_name = ""
                for record in tt['name'].names:
                    if record.nameID == 1:
                        try:
                            extracted_name = record.toUnicode()
                            break
                        except Exception:
                            pass
                st.session_state.font_name = extracted_name
            except Exception:
                st.session_state.font_name = ""
        else:
            st.session_state.font_name = ""

        st.session_state.file_bytes = file_bytes
        base = (raw_name or "translated").replace(".pptx", "")
        st.session_state.file_name = f"{base}_EN.pptx"
        st.session_state.text_units = text_units
        st.session_state.classification_done = False
        st.session_state.stage = "classify"
        st.rerun()

# ── Stage: classify ───────────────────────────────────────────────────────────

elif st.session_state.stage == "classify":
    st.header("1단계: 텍스트 분류")

    if not st.session_state.classification_done:
        with st.spinner("AI가 텍스트를 분류하는 중..."):
            try:
                classified = classify_text_units(
                    st.session_state.text_units, _build_context()
                )
                st.session_state.classified_units = classified
                st.session_state.classification_done = True
            except Exception as e:
                st.error(f"분류 오류: {e}")
                st.stop()
        st.rerun()

    st.info("AI 분류 결과를 확인하고 필요하면 수정하세요. 완료 후 '확인' 버튼을 누르세요.")

    units = st.session_state.classified_units

    # Group by slide
    from collections import defaultdict as _dd
    slide_groups = _dd(list)
    for i, u in enumerate(units):
        slide_groups[u["slide_idx"]].append((i, u))

    rerun_needed = False
    for slide_idx in sorted(slide_groups.keys()):
        items = slide_groups[slide_idx]
        hcol, bcol1, bcol2 = st.columns([4, 1, 1])
        with hcol:
            st.markdown(f"**슬라이드 {slide_idx + 1}**")
        with bcol1:
            if st.button("전체 발표용", key=f"all_pres_{slide_idx}", use_container_width=True):
                for i, _ in items:
                    st.session_state.classified_units[i]["category"] = "presentation"
                rerun_needed = True
        with bcol2:
            if st.button("전체 카피", key=f"all_copy_{slide_idx}", use_container_width=True):
                for i, _ in items:
                    st.session_state.classified_units[i]["category"] = "copy"
                rerun_needed = True

        for i, u in items:
            cat = u["category"]
            tcol, bcol = st.columns([6, 1])
            with tcol:
                st.markdown(u["ko_text"])
            with bcol:
                btn_label = "📢 카피" if cat == "copy" else "📊 발표용"
                btn_type = "primary" if cat == "copy" else "secondary"
                if st.button(btn_label, key=f"tog_{u['id']}", type=btn_type, use_container_width=True):
                    st.session_state.classified_units[i]["category"] = (
                        "presentation" if cat == "copy" else "copy"
                    )
                    rerun_needed = True
        st.divider()

    if rerun_needed:
        st.rerun()

    col_back, col_confirm = st.columns([1, 5])
    with col_back:
        if st.button("← 이전"):
            st.session_state.stage = "upload"
            st.rerun()
    with col_confirm:
        if st.button("확인 →", type="primary"):
            pres = [u for u in st.session_state.classified_units if u["category"] == "presentation"]
            copy = [u for u in st.session_state.classified_units if u["category"] == "copy"]
            st.session_state.presentation_units = pres
            st.session_state.copy_units = copy
            st.session_state.translations_loaded = False
            st.session_state.copy_options_loaded = False
            st.session_state.current_copy_idx = 0
            st.session_state.chat_history_2a = []
            st.session_state.chat_history_2b = []

            if pres:
                st.session_state.stage = "review_2a"
            elif copy:
                st.session_state.stage = "review_2b"
            else:
                st.error("분류된 텍스트가 없습니다.")
                st.stop()
            st.rerun()

# ── Stage: review_2a ──────────────────────────────────────────────────────────

elif st.session_state.stage == "review_2a":
    st.header("2단계-A: 발표용 텍스트 번역 감수")

    pres_units = st.session_state.presentation_units

    if not st.session_state.translations_loaded:
        with st.spinner("AI가 발표용 텍스트를 일괄 번역하는 중..."):
            try:
                trans = translate_presentation_texts(pres_units, _build_context())
                st.session_state.presentation_translations = trans
                st.session_state.translations_loaded = True
            except Exception as e:
                st.error(f"번역 오류: {e}")
                st.stop()
        st.rerun()

    # Slide-grouped translation cards
    trans = st.session_state.presentation_translations

    from collections import defaultdict as _dd2
    slide_groups_2a = _dd2(list)
    for u in pres_units:
        slide_groups_2a[u["slide_idx"]].append(u)

    for slide_idx in sorted(slide_groups_2a.keys()):
        items = slide_groups_2a[slide_idx]
        st.markdown(f"**슬라이드 {slide_idx + 1}**")
        for u in items:
            item = trans.get(u["id"], {})
            en_text = item.get("en_text", "")
            notes = item.get("notes", "")
            clarification = item.get("clarification", "")
            ko_col, en_col = st.columns(2)
            with ko_col:
                st.markdown(
                    f"<div style='padding:10px 14px;border-radius:6px;"
                    f"border:1px solid rgba(128,128,128,0.3);font-size:14px;"
                    f"line-height:1.6;'>{u['ko_text']}</div>",
                    unsafe_allow_html=True,
                )
            with en_col:
                st.markdown(
                    f"<div style='padding:10px 14px;border-radius:6px;"
                    f"border:1px solid rgba(128,128,128,0.3);font-size:14px;"
                    f"line-height:1.6;'>{en_text if en_text else '—'}</div>",
                    unsafe_allow_html=True,
                )
            if clarification:
                st.caption(f"💬 해석 가정: {clarification}")
            if notes:
                st.caption(f"📝 번역 노트: {notes}")
        st.divider()

    # Chat history display
    for msg in st.session_state.chat_history_2a:
        with st.chat_message(msg["role"]):
            st.write(msg["content"])

    # Chat input
    st.caption("수정 지시사항을 입력하면 AI가 번역을 업데이트합니다. (예: \"모든 동사를 능동형으로\")")
    user_msg = st.chat_input("번역 수정 지시사항 입력...")
    if user_msg:
        st.session_state.chat_history_2a.append({"role": "user", "content": user_msg})
        with st.spinner("번역 수정 중..."):
            try:
                updated = chat_modify_presentation(
                    pres_units,
                    st.session_state.presentation_translations,
                    user_msg,
                    _build_context(),
                )
                st.session_state.presentation_translations = updated
                st.session_state.chat_history_2a.append(
                    {"role": "assistant", "content": "번역을 업데이트했습니다."}
                )
            except Exception as e:
                st.session_state.chat_history_2a.append(
                    {"role": "assistant", "content": f"오류: {e}"}
                )
        st.rerun()

    col_back, _, col_next = st.columns([1, 4, 1])
    with col_back:
        if st.button("← 분류로"):
            st.session_state.stage = "classify"
            st.rerun()
    with col_next:
        next_label = "카피 선택 →" if st.session_state.copy_units else "다운로드 →"
        if st.button(next_label, type="primary"):
            if st.session_state.copy_units:
                st.session_state.stage = "review_2b"
            else:
                st.session_state.stage = "download"
            st.rerun()

# ── Stage: review_2b ──────────────────────────────────────────────────────────

elif st.session_state.stage == "review_2b":
    st.header("2단계-B: 광고 카피 선택 및 감수")

    copy_units = st.session_state.copy_units

    if not st.session_state.copy_options_loaded:
        with st.spinner("AI가 카피 옵션을 생성하는 중... (뉘앙스 분석 포함, 잠시만 기다려주세요)"):
            try:
                opts = generate_copy_options(copy_units, _build_context())
                st.session_state.copy_options = opts
                # Default selection: first (creative) option
                st.session_state.copy_selections = {
                    u["id"]: opts.get(u["id"], {}).get("options", [""])[0]
                    for u in copy_units
                }
                st.session_state.copy_options_loaded = True
            except Exception as e:
                st.error(f"카피 생성 오류: {e}")
                st.stop()
        st.rerun()

    idx = st.session_state.current_copy_idx
    total = len(copy_units)

    if idx >= total:
        idx = total - 1
        st.session_state.current_copy_idx = idx

    unit = copy_units[idx]
    unit_data = st.session_state.copy_options.get(unit["id"], {})
    options = unit_data.get("options", ["", "", ""])
    notes = unit_data.get("notes", "")
    recommendation = unit_data.get("recommendation", "")
    cultural_flag = unit_data.get("cultural_flag", "")
    clarification = unit_data.get("clarification", "")
    current_selection = st.session_state.copy_selections.get(unit["id"], options[0] if options else "")

    # Progress
    st.progress((idx) / total, text=f"{idx + 1} / {total} 라인")

    # Left (Korean) | Right (3 options)
    col_ko, col_en = st.columns(2)
    with col_ko:
        st.markdown("**한국어 원문**")
        st.markdown(
            f"""<div style="padding:16px;border-radius:8px;font-size:15px;line-height:1.6;border:1px solid rgba(128,128,128,0.3);">
            {unit['ko_text']}</div>""",
            unsafe_allow_html=True,
        )

        # Clarification flag (if AI had to assume meaning)
        if clarification:
            st.warning(f"**해석 가정:** {clarification}")

        # Cultural flag
        if cultural_flag:
            st.info(f"**문화적 참고:** {cultural_flag}")

    with col_en:
        st.markdown("**영어 옵션 선택**")

        OPTION_LABELS = [
            "① Creative — 의역 (feel·rhythm·impact 우선)",
            "② Balanced — 균형 (의미 + 자연스러운 영어)",
            "③ Faithful — 직역 (원문 의미 충실)",
        ]

        try:
            sel_idx = options.index(current_selection)
        except ValueError:
            sel_idx = None

        if sel_idx is not None:
            chosen_idx = st.radio(
                "옵션",
                options=list(range(len(options))),
                index=sel_idx,
                format_func=lambda i: f"{OPTION_LABELS[i]}\n\n{options[i]}",
                label_visibility="collapsed",
                key=f"radio_{unit['id']}",
            )
            chosen = options[chosen_idx]
            st.session_state.copy_selections[unit["id"]] = chosen
        else:
            st.success(f"**수정된 카피:** {current_selection}")
            if st.button("원래 옵션으로 돌아가기"):
                st.session_state.copy_selections[unit["id"]] = options[0]
                st.session_state.chat_history_2b = []
                st.rerun()

    # Translator notes (always shown)
    if notes or recommendation:
        with st.expander("📝 번역 노트 (뉘앙스·라임·추천)", expanded=True):
            if notes:
                st.markdown(f"**원문 분석**\n\n{notes}")
            if recommendation:
                st.markdown(f"**추천**\n\n{recommendation}")

    st.divider()

    # Chat history for current line
    for msg in st.session_state.chat_history_2b:
        with st.chat_message(msg["role"]):
            st.write(msg["content"])

    st.caption("선택한 옵션을 추가로 수정하려면 지시사항을 입력하세요.")
    user_msg = st.chat_input(f"카피 수정 지시사항 입력... ({idx + 1}/{total})")
    if user_msg:
        st.session_state.chat_history_2b.append({"role": "user", "content": user_msg})
        sel = st.session_state.copy_selections.get(unit["id"], options[0])
        with st.spinner("카피 수정 중..."):
            try:
                refined = chat_refine_copy(
                    unit["ko_text"], sel, user_msg, _build_context()
                )
                st.session_state.copy_selections[unit["id"]] = refined
                st.session_state.chat_history_2b.append(
                    {"role": "assistant", "content": f"수정 완료: {refined}"}
                )
            except Exception as e:
                st.session_state.chat_history_2b.append(
                    {"role": "assistant", "content": f"오류: {e}"}
                )
        st.rerun()

    # Navigation
    col_back, col_prev, col_next = st.columns([1, 1, 1])
    with col_back:
        if st.session_state.presentation_units:
            if st.button("← 발표용으로"):
                st.session_state.stage = "review_2a"
                st.rerun()
    with col_prev:
        if idx > 0:
            if st.button("← 이전 라인"):
                st.session_state.current_copy_idx -= 1
                st.session_state.chat_history_2b = []
                st.rerun()
    with col_next:
        if idx < total - 1:
            if st.button("다음 라인 →"):
                st.session_state.current_copy_idx += 1
                st.session_state.chat_history_2b = []
                st.rerun()
        else:
            if st.button("완료 →", type="primary"):
                st.session_state.stage = "download"
                st.rerun()

# ── Stage: download ───────────────────────────────────────────────────────────

elif st.session_state.stage == "download":
    st.header("번역 완료 — 파일 다운로드")

    if st.session_state.output_bytes is None:
        with st.spinner("번역본을 PPTX에 적용하는 중..."):
            # presentation_translations stores {id: {en_text, notes, ...}}; extract just en_text
            pres_trans = {
                uid: v.get("en_text", "")
                for uid, v in st.session_state.presentation_translations.items()
            }
            all_translations = {
                **pres_trans,
                **st.session_state.copy_selections,
            }
            try:
                out = apply_translations(
                    st.session_state.file_bytes,
                    all_translations,
                    font_name=st.session_state.font_name,
                )
                st.session_state.output_bytes = out
            except Exception as e:
                st.error(f"파일 생성 오류: {e}")
                st.stop()
        st.rerun()

    st.success("번역이 완료되었습니다!")

    # Summary
    n_pres = len(st.session_state.presentation_units)
    n_copy = len(st.session_state.copy_units)
    st.markdown(
        f"- 발표용 텍스트: **{n_pres}개** 번역\n"
        f"- 광고 카피: **{n_copy}개** 트랜스크리에이션"
    )

    st.download_button(
        label="영문 PPT 다운로드",
        data=st.session_state.output_bytes,
        file_name=st.session_state.file_name,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
    )

    if st.button("처음으로 돌아가기"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
