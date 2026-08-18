import streamlit as st
import pandas as pd
import io

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

if st.session_state.stage == "upload":
    st.header("PPT 파일 업로드 및 캠페인 브리프")

    # ── 1. PPTX 파일 ──────────────────────────────────────────────────────────
    uploaded = st.file_uploader("한국어 PPTX 파일 선택 *", type=["pptx"])

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

    # ── 5. 고유명사 (번역 유지) ───────────────────────────────────────────────
    st.subheader("고유명사 / 번역하지 않을 단어")
    st.caption("영어로도 그대로 쓸 브랜드명, 인명, 제품명 등을 쉼표로 구분해 입력하세요.")
    glossary = st.text_input(
        "고유명사",
        placeholder="예: ChatGPT, POSCO, K-Beauty",
        value=st.session_state.glossary,
        label_visibility="collapsed",
    )

    st.divider()

    if st.button("분석 시작", type="primary", disabled=uploaded is None):
        file_bytes = uploaded.read()
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

        st.session_state.file_bytes = file_bytes
        st.session_state.file_name = uploaded.name.replace(".pptx", "_EN.pptx")
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
    df = pd.DataFrame(
        [
            {
                "슬라이드": u["slide_idx"] + 1,
                "한국어 텍스트": u["ko_text"],
                "분류": "카피" if u["category"] == "copy" else "발표용",
            }
            for u in units
        ]
    )

    edited_df = st.data_editor(
        df,
        column_config={
            "슬라이드": st.column_config.NumberColumn(disabled=True, width="small"),
            "한국어 텍스트": st.column_config.TextColumn(disabled=True, width="large"),
            "분류": st.column_config.SelectboxColumn(
                options=["발표용", "카피"], required=True, width="medium"
            ),
        },
        hide_index=True,
        use_container_width=True,
        key="classify_editor",
    )

    col_back, col_confirm = st.columns([1, 5])
    with col_back:
        if st.button("← 이전"):
            st.session_state.stage = "upload"
            st.rerun()
    with col_confirm:
        if st.button("확인 →", type="primary"):
            for i, row in edited_df.iterrows():
                st.session_state.classified_units[i]["category"] = (
                    "copy" if row["분류"] == "카피" else "presentation"
                )

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

    # Left-right comparison table
    st.subheader("번역 결과 (좌: 한국어 / 우: 영어)")
    trans = st.session_state.presentation_translations
    rows = []
    for u in pres_units:
        item = trans.get(u["id"], {})
        rows.append({
            "슬라이드": u["slide_idx"] + 1,
            "한국어": u["ko_text"],
            "영어 (번역)": item.get("en_text", ""),
        })
    st.dataframe(
        pd.DataFrame(rows),
        column_config={
            "슬라이드": st.column_config.NumberColumn(width="small"),
            "한국어": st.column_config.TextColumn(width="large"),
            "영어 (번역)": st.column_config.TextColumn(width="large"),
        },
        hide_index=True,
        use_container_width=True,
    )

    # Show translator notes if any
    flagged = [
        (u, trans.get(u["id"], {}))
        for u in pres_units
        if trans.get(u["id"], {}).get("notes") or trans.get(u["id"], {}).get("clarification")
    ]
    if flagged:
        with st.expander(f"📝 번역 노트 ({len(flagged)}개 항목)", expanded=False):
            for u, item in flagged:
                st.markdown(f"**슬라이드 {u['slide_idx'] + 1}** — {u['ko_text']}")
                st.markdown(f"> {item.get('en_text', '')}")
                if item.get("clarification"):
                    st.warning(f"**해석 가정:** {item['clarification']}")
                if item.get("notes"):
                    st.info(f"**번역 노트:** {item['notes']}")
                st.markdown("---")

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
            f"""<div style="background:#f0f2f6;padding:16px;border-radius:8px;font-size:15px;line-height:1.6;">
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
                out = apply_translations(st.session_state.file_bytes, all_translations)
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
