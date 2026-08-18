import streamlit as st
import pandas as pd
import io

from pptx_utils import extract_text_units, apply_translations
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
        "glossary": "",
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

# ── Stage: upload ─────────────────────────────────────────────────────────────

if st.session_state.stage == "upload":
    st.header("PPT 파일 업로드 및 사전 설정")

    uploaded = st.file_uploader("한국어 PPTX 파일 선택", type=["pptx"])
    glossary = st.text_area(
        "고유명사 / 유지 단어 (Glossary)",
        placeholder="예: OpenAI, 삼성전자, ChatGPT → 번역 시 그대로 유지됩니다.",
        height=80,
    )

    if st.button("분석 시작", type="primary", disabled=uploaded is None):
        file_bytes = uploaded.read()
        with st.spinner("텍스트 파싱 중..."):
            text_units = extract_text_units(file_bytes)

        if not text_units:
            st.error("번역 가능한 텍스트를 찾지 못했습니다. 파일을 확인해 주세요.")
            st.stop()

        st.session_state.file_bytes = file_bytes
        st.session_state.file_name = uploaded.name.replace(".pptx", "_EN.pptx")
        st.session_state.glossary = glossary
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
                    st.session_state.text_units, st.session_state.glossary
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
                trans = translate_presentation_texts(pres_units, st.session_state.glossary)
                st.session_state.presentation_translations = trans
                st.session_state.translations_loaded = True
            except Exception as e:
                st.error(f"번역 오류: {e}")
                st.stop()
        st.rerun()

    # Left-right comparison table
    st.subheader("번역 결과 (좌: 한국어 / 우: 영어)")
    rows = []
    for u in pres_units:
        rows.append({
            "슬라이드": u["slide_idx"] + 1,
            "한국어": u["ko_text"],
            "영어 (번역)": st.session_state.presentation_translations.get(u["id"], ""),
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
                    st.session_state.glossary,
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
        with st.spinner("AI가 카피 옵션을 생성하는 중..."):
            try:
                opts = generate_copy_options(copy_units, st.session_state.glossary)
                st.session_state.copy_options = opts
                # Default selection: first option
                st.session_state.copy_selections = {
                    u["id"]: opts.get(u["id"], ["", "", ""])[0] for u in copy_units
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
    options = st.session_state.copy_options.get(unit["id"], ["", "", ""])
    current_selection = st.session_state.copy_selections.get(unit["id"], options[0] if options else "")

    # Progress
    st.progress((idx) / total, text=f"{idx + 1} / {total} 라인")

    # Left (Korean) | Right (3 options)
    col_ko, col_en = st.columns(2)
    with col_ko:
        st.markdown("**한국어 원문**")
        st.markdown(
            f"""<div style="background:#f0f2f6;padding:16px;border-radius:8px;font-size:15px;">
            {unit['ko_text']}</div>""",
            unsafe_allow_html=True,
        )

    with col_en:
        st.markdown("**영어 옵션 선택**")
        # Find index of current selection (may be a custom refined version)
        try:
            sel_idx = options.index(current_selection)
        except ValueError:
            sel_idx = None

        if sel_idx is not None:
            chosen = st.radio(
                "옵션",
                options=options,
                index=sel_idx,
                label_visibility="collapsed",
                key=f"radio_{unit['id']}",
            )
            st.session_state.copy_selections[unit["id"]] = chosen
        else:
            # Custom refined text — show as highlighted + option to reset
            st.success(f"**수정된 카피:** {current_selection}")
            if st.button("원래 옵션으로 돌아가기"):
                st.session_state.copy_selections[unit["id"]] = options[0]
                # Reset chat history for this line
                st.session_state.chat_history_2b = []
                st.rerun()

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
                    unit["ko_text"], sel, user_msg, st.session_state.glossary
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
            all_translations = {
                **st.session_state.presentation_translations,
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
