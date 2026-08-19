import streamlit as st
import pandas as pd
import io
import re
import html as _html
import requests
import base64
from collections import defaultdict

from pptx_utils import extract_text_units, apply_translations, extract_reference_texts, render_slides_to_images
from ai_utils import (
    classify_text_units,
    translate_presentation_texts,
    generate_copy_options,
    chat_modify_presentation,
    chat_refine_copy,
)

st.set_page_config(
    page_title="광고주 제안 문서 영문 번역 시스템",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Design system CSS ─────────────────────────────────────────────────────────
st.markdown("""<style>
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable.min.css');

#MainMenu, footer, .stDeployButton { display: none !important; }
[data-testid="stHeader"] { display: none !important; }
[data-testid="stSidebar"] { display: none !important; }
.block-container { padding: 0 !important; max-width: 100% !important; }

:root {
  --cp: #0C2790;
  --cy: #F2CC54;
  --cbg: #F5F6F8;
  --ct: #101828;
  --cm: #667085;
  --cb: #E4E7EC;
  --cw: #ffffff;
  --cbl: #E8EDFB;
  --r: 12px;
  --font: 'Pretendard Variable', -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif;
}
body, .stApp, .main { background: var(--cbg) !important; font-family: var(--font) !important; }
section[data-testid="stMain"] { background: var(--cbg) !important; }
* { box-sizing: border-box; }

/* Data editor — keep grid light */
[data-testid="stDataFrameResizable"] {
  border: 1px solid var(--cb) !important; border-radius: 8px !important;
  overflow: hidden !important; background: var(--cw) !important;
}
.glideDataEditor, .dvn-scroller { background: var(--cw) !important; }

/* File uploader dropzone — light gray, not black */
[data-testid="stFileUploadDropzone"] {
  background: var(--cbg) !important;
  border: 1.5px dashed var(--cb) !important;
  border-radius: 8px !important;
}
[data-testid="stFileUploadDropzone"] section { background: transparent !important; }
[data-testid="stFileUploadDropzone"] p,
[data-testid="stFileUploadDropzone"] span { color: var(--cm) !important; }

/* Streamlit container wrappers — keep transparent */
[data-testid="stVerticalBlock"] > [data-testid="stVerticalBlock"] { background: transparent !important; }
[data-testid="column"] { background: transparent !important; }

/* Top bar */
.topbar {
  display: flex; align-items: center; gap: 12px;
  padding: 0 32px; height: 64px;
  background: var(--cw); border-bottom: 1px solid var(--cb);
}
.topbar-logo {
  width: 36px; height: 36px; border-radius: 8px; background: var(--cp);
  display: flex; align-items: center; justify-content: center; flex-shrink: 0;
}
.topbar-title { font-size: 15px; font-weight: 600; color: var(--ct); }

/* Step rail */
.steprail {
  display: flex; align-items: stretch; border-bottom: 1px solid var(--cb);
  background: var(--cw); padding: 0 32px; height: 52px;
}
.step {
  display: flex; align-items: center; gap: 8px; padding: 0 18px;
  font-size: 13px; font-weight: 500; color: var(--cm);
  position: relative; white-space: nowrap;
}
.step.active { color: var(--ct); }
.step.active::after {
  content: ''; position: absolute; bottom: 0; left: 0; right: 0;
  height: 2.5px; background: var(--cp); border-radius: 2px 2px 0 0;
}
.step-n {
  width: 24px; height: 24px; border-radius: 50%;
  display: flex; align-items: center; justify-content: center;
  font-size: 12px; font-weight: 700; flex-shrink: 0;
  border: 1.5px solid var(--cb); background: var(--cbg); color: var(--cm);
}
.step.active .step-n { background: var(--cp); color: #fff; border-color: var(--cp); }
.step.done .step-n  { background: var(--cbl); color: var(--cp); border-color: #B8C4ED; }
.step.done { color: var(--cm); }

/* Card */
.card {
  background: var(--cw); border: 1px solid var(--cb);
  border-radius: var(--r); padding: 24px; margin-bottom: 16px;
}
.card-title { font-size: 14px; font-weight: 600; color: var(--ct); margin-bottom: 4px; }
.card-sub { font-size: 12.5px; color: var(--cm); margin-bottom: 14px; }
.field-label { font-size: 12.5px; color: var(--cm); font-weight: 500; margin-bottom: 6px; }

/* Slide bezel */
.bezel { background: #1B1B2E; border-radius: 14px; padding: 18px; margin-bottom: 0; }
.bezel img { border-radius: 6px; width: 100%; display: block; }
.bezel-placeholder {
  border-radius: 6px; background: #2D2D42; width: 100%; aspect-ratio: 16/9;
  display: flex; align-items: center; justify-content: center;
  color: rgba(255,255,255,0.35); font-size: 14px;
}
.bezel-caption { color: rgba(255,255,255,0.72); font-size: 13px; margin-top: 12px; text-align: center; }

/* Note / translation boxes */
.notebox {
  background: #F2F4F7; border-radius: 8px; padding: 14px 16px;
  display: flex; gap: 12px; align-items: flex-start;
  font-size: 13.5px; line-height: 1.65; color: var(--ct); margin: 12px 0;
}
.notebox-icon { font-size: 18px; flex-shrink: 0; margin-top: 1px; }
.transbox {
  background: var(--cbl); border-radius: 8px; padding: 22px 28px;
  text-align: center; font-size: 22px; font-weight: 700;
  color: var(--cp); line-height: 1.4; margin: 12px 0;
}

/* Copy option rows */
.copyrow {
  display: flex; align-items: center; gap: 16px;
  padding: 14px 18px; border-radius: 10px; border: 1.5px solid var(--cb);
  margin-bottom: 8px; background: var(--cw);
}
.copyrow.sel { background: var(--cbl); border-color: var(--cp); }
.cr-label { flex: 0 0 130px; }
.cr-lname { font-size: 13px; font-weight: 600; color: var(--ct); }
.cr-lsub { font-size: 11.5px; color: var(--cm); margin-top: 1px; }
.cr-text { flex: 1; font-size: 15px; font-weight: 600; color: var(--ct); }
.cr-star { font-size: 20px; flex-shrink: 0; }

/* Recommendation */
.recbox {
  background: var(--cbl); border-radius: 8px; padding: 14px 18px;
  display: flex; gap: 10px; font-size: 13px; line-height: 1.65;
  color: var(--ct); margin-bottom: 12px;
}

/* Tags */
.tag-p { display:inline-flex;align-items:center;gap:4px;padding:3px 10px;
  background:var(--cy);color:#101828;border-radius:20px;font-size:12px;font-weight:600; }
.tag-c { display:inline-flex;align-items:center;gap:4px;padding:3px 10px;
  background:var(--cp);color:white;border-radius:20px;font-size:12px;font-weight:600; }

/* Legend */
.legend-row { display:flex;align-items:center;gap:8px;margin-bottom:10px;font-size:13px;color:var(--ct); }
.dot-y { width:12px;height:12px;border-radius:50%;background:var(--cy);flex-shrink:0; }
.dot-b { width:12px;height:12px;border-radius:50%;background:var(--cp);flex-shrink:0; }

/* Slide text item in classify right panel */
.cl-item {
  display:flex;align-items:flex-start;justify-content:space-between;gap:10px;
  padding:10px 14px;border-radius:8px;background:var(--cbg);
  border:1px solid var(--cb);margin-bottom:8px;
}
.cl-item-text { font-size:13px;color:var(--ct);line-height:1.5;flex:1; }

/* Button overrides */
div.stButton > button { border-radius:8px !important; font-weight:500 !important; }
div.stButton > button[kind="primary"] {
  background:var(--cp) !important; border-color:var(--cp) !important; color:#fff !important;
}

/* ── Input fields: force light mode ────────────────────────────────────────── */
/* text_input */
.stTextInput > div > div > input {
  background: #ffffff !important;
  color: #101828 !important;
  border: 1px solid #E4E7EC !important;
  border-radius: 8px !important;
  font-size: 14px !important;
  box-shadow: none !important;
}
.stTextInput > div > div > input::placeholder { color: #98A2B3 !important; }
.stTextInput > div > div > input:focus {
  border-color: #0C2790 !important;
  box-shadow: 0 0 0 3px rgba(12,39,144,0.12) !important;
}
/* text_area */
.stTextArea > div > div > textarea {
  background: #ffffff !important;
  color: #101828 !important;
  border: 1px solid #E4E7EC !important;
  border-radius: 8px !important;
  font-size: 14px !important;
  box-shadow: none !important;
}
.stTextArea > div > div > textarea::placeholder { color: #98A2B3 !important; }
.stTextArea > div > div > textarea:focus {
  border-color: #0C2790 !important;
  box-shadow: 0 0 0 3px rgba(12,39,144,0.12) !important;
}
/* select box */
.stSelectbox > div > div > div {
  background: #ffffff !important;
  color: #101828 !important;
  border: 1px solid #E4E7EC !important;
  border-radius: 8px !important;
}
/* generic Streamlit input wrapper */
[data-baseweb="input"] { background: #ffffff !important; }
[data-baseweb="input"] input { background: transparent !important; color: #101828 !important; }
[data-baseweb="textarea"] { background: #ffffff !important; }
[data-baseweb="textarea"] textarea { background: transparent !important; color: #101828 !important; }
/* label text above inputs */
.stTextInput label, .stTextArea label, .stSelectbox label {
  color: #667085 !important; font-size: 12.5px !important; font-weight: 500 !important;
}

/* Nav row */
.navrow { display:flex;align-items:center;justify-content:space-between;padding:16px 0;gap:10px; }

/* Sticky left column — slide preview stays fixed while right scrolls */
[data-testid="stHorizontalBlock"] {
  align-items: flex-start !important;
}
.sticky-slide {
  position: sticky;
  top: 116px;   /* topbar 64px + steprail 52px */
}
.scroll-panel {
  max-height: calc(100vh - 180px);
  overflow-y: auto;
  padding-right: 4px;
}
</style>""", unsafe_allow_html=True)


# ── Session state ─────────────────────────────────────────────────────────────
def _init():
    defaults = {
        "stage": "upload",
        "file_bytes": None,
        "file_name": "translated.pptx",
        "brand_name_ko": "",
        "brand_name_en": "",
        "key_phrases": [],
        "ref_pptx_texts": [],
        "glossary": "",
        "font_name": "",
        "text_units": [],
        "classified_units": [],
        "classification_done": False,
        "presentation_units": [],
        "copy_units": [],
        "presentation_translations": {},
        "translations_loaded": False,
        "copy_options": {},
        "copy_selections": {},
        "copy_options_loaded": False,
        "current_pres_idx": 0,
        "current_copy_idx": 0,
        "chat_history_2b": [],
        "output_bytes": None,
        "slide_images": [],
        "slide_count": 0,
        "active_classify_slide": 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init()


# ── Google Drive helpers ──────────────────────────────────────────────────────
def _gdrive_file_id(url: str) -> str | None:
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
    session = requests.Session()
    errors = []
    if is_slides:
        url = f"https://docs.google.com/presentation/d/{file_id}/export/pptx"
        try:
            resp = session.get(url, timeout=300)
            if resp.status_code == 200 and len(resp.content) > 1000:
                return resp.content
            errors.append(f"Slides export: HTTP {resp.status_code}")
        except Exception as e:
            errors.append(f"Slides export: {e}")
    url = f"https://drive.usercontent.google.com/download?id={file_id}&export=download&authuser=0&confirm=t"
    try:
        resp = session.get(url, stream=True, timeout=300)
        if resp.status_code == 200 and len(resp.content) > 1000:
            return resp.content
        errors.append(f"usercontent: HTTP {resp.status_code}")
    except Exception as e:
        errors.append(f"usercontent: {e}")
    if not is_slides:
        url = f"https://docs.google.com/presentation/d/{file_id}/export/pptx"
        try:
            resp = session.get(url, timeout=300)
            if resp.status_code == 200 and len(resp.content) > 1000:
                return resp.content
            errors.append(f"Slides fallback: HTTP {resp.status_code}")
        except Exception as e:
            errors.append(f"Slides fallback: {e}")
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


def _build_context() -> str:
    parts = []
    brand_ko = st.session_state.brand_name_ko.strip()
    brand_en = st.session_state.brand_name_en.strip()
    if brand_en or brand_ko:
        label = f"'{brand_en}' (EN) / '{brand_ko}' (KO)" if brand_en and brand_ko else brand_en or brand_ko
        parts.append(f"Brand name: {label}. Always use '{brand_en or brand_ko}' in English translations.")
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
        parts.append(f"Style & terminology reference (from a previous approved English translation):\n{sample}")
    return "\n\n".join(parts)


def _slide_img_html(slide_idx: int) -> str:
    imgs = st.session_state.slide_images
    if imgs and slide_idx < len(imgs):
        b64 = base64.b64encode(imgs[slide_idx]).decode()
        return f'<img src="data:image/png;base64,{b64}" style="border-radius:6px;width:100%;display:block;" />'
    return f'<div class="bezel-placeholder">슬라이드 {slide_idx + 1}</div>'

def _thumb_html(slide_idx: int, active: bool) -> str:
    border = "2px solid #0C2790" if active else "2px solid transparent"
    imgs = st.session_state.slide_images
    inner = _slide_img_html(slide_idx)
    return (
        f'<div style="border:{border};border-radius:6px;overflow:hidden;margin-bottom:4px;cursor:pointer;">'
        f'{inner}'
        f'<div style="text-align:center;font-size:11px;color:#667085;padding:2px 0 4px;background:#fff;">슬라이드 {slide_idx+1}</div>'
        f'</div>'
    )


# ── Chrome: top bar + step rail ───────────────────────────────────────────────
STAGES = ["upload", "classify", "review_2a", "review_2b", "download"]
STAGE_LABELS = ["업로드", "분류", "발표용 감수", "카피 선택", "다운로드"]

def _render_chrome():
    curr_idx = STAGES.index(st.session_state.stage)
    st.markdown(
        '<div class="topbar">'
        '<div class="topbar-logo">'
        '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">'
        '<rect x="2" y="3" width="20" height="15" rx="2" stroke="white" stroke-width="2" fill="none"/>'
        '<path d="M8 22h8M12 18v4" stroke="white" stroke-width="2" stroke-linecap="round"/>'
        '</svg></div>'
        '<span class="topbar-title">광고주 제안 문서 영문 번역 시스템</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    items = ""
    for i, label in enumerate(STAGE_LABELS):
        if i < curr_idx:
            cls = "step done"
        elif i == curr_idx:
            cls = "step active"
        else:
            cls = "step"
        items += f'<div class="{cls}"><span class="step-n">{i+1}</span>{label}</div>'
    st.markdown(f'<div class="steprail">{items}</div>', unsafe_allow_html=True)

_render_chrome()


# ── Stage: upload ─────────────────────────────────────────────────────────────
if st.session_state.stage == "upload":
    st.markdown('<div class="page">', unsafe_allow_html=True)
    st.markdown('<div class="page-title">PPT 파일 업로드 및 캠페인 브리프</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-sub">번역할 한국어 PPTX 파일과 캠페인 정보를 입력하세요.</div>', unsafe_allow_html=True)

    # Card 1: PPTX file
    st.markdown('<div class="card"><div class="card-title">한국어 PPTX 파일</div>', unsafe_allow_html=True)
    st.markdown('<div class="field-label">☁ Google Drive 링크</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([5, 1])
    with c1:
        drive_url = st.text_input(
            "drive_url", placeholder="https://drive.google.com/file/d/...",
            label_visibility="collapsed", key="up_drive_url",
        )
    with c2:
        fetch_clicked = st.button("가져오기", key="btn_fetch_main", use_container_width=True, type="primary")
    if drive_url.strip() and not _gdrive_file_id(drive_url.strip()):
        st.warning("올바른 Google Drive 링크가 아닙니다.")
    st.markdown('</div>', unsafe_allow_html=True)

    # Card 2: Brand name
    st.markdown('<div class="card"><div class="card-title">브랜드명</div>', unsafe_allow_html=True)
    c_ko, c_en = st.columns(2)
    with c_ko:
        st.markdown('<div class="field-label">한국어</div>', unsafe_allow_html=True)
        brand_ko = st.text_input("brand_ko", placeholder="예: 삼성전자",
            value=st.session_state.brand_name_ko, label_visibility="collapsed")
    with c_en:
        st.markdown('<div class="field-label">영어</div>', unsafe_allow_html=True)
        brand_en = st.text_input("brand_en", placeholder="e.g. Samsung Electronics",
            value=st.session_state.brand_name_en, label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

    # Card 3: Term mapping
    st.markdown(
        '<div class="card"><div class="card-title">주요 용어 매핑 (한국어 → 영어)</div>'
        '<div class="card-sub">자주 쓰는 표현의 선호 번역을 지정합니다. 행 추가 버튼으로 한 줄을 추가하세요.</div>',
        unsafe_allow_html=True,
    )
    init_kp = st.session_state.key_phrases if st.session_state.key_phrases else [{"한국어": "", "영어": ""}]
    edited_kp = st.data_editor(
        pd.DataFrame(init_kp),
        column_config={
            "한국어": st.column_config.TextColumn("한국어 표현", width="large"),
            "영어": st.column_config.TextColumn("영어 번역 (선호)", width="large"),
        },
        num_rows="dynamic", hide_index=True, use_container_width=True, key="kp_editor",
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # Card 4: Reference PPTX (optional)
    st.markdown(
        '<div class="card"><div class="card-title">이전 번역본 참고 (선택)</div>'
        '<div class="card-sub">기존 영문 PPT를 올리면 용어·문체를 참고해 일관성을 유지합니다.</div>'
        '<div class="field-label">영문 참고 PPTX 링크 (선택)</div>',
        unsafe_allow_html=True,
    )
    c3, c4 = st.columns([5, 1])
    with c3:
        ref_drive_url = st.text_input(
            "ref_url", placeholder="https://drive.google.com/file/d/...",
            label_visibility="collapsed", key="up_ref_url",
        )
    with c4:
        st.button("가져오기", key="btn_fetch_ref", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Card 5: Font (optional)
    st.markdown(
        '<div class="card"><div class="card-title">영어 폰트 (선택) ✏️</div>'
        '<div class="card-sub">번역된 텍스트에 적용할 TTF/OTF 폰트 파일을 업로드하세요. (해당 폰트가 PPT를 열 PC에 설치되어 있어야 합니다)</div>',
        unsafe_allow_html=True,
    )
    font_file = st.file_uploader("폰트 파일 (선택, TTF/OTF)", type=["ttf", "otf"], key="font_uploader",
                                  label_visibility="collapsed")
    if font_file:
        st.caption(f"업로드: {font_file.name}")
    st.markdown('</div>', unsafe_allow_html=True)

    # Card 6: Proper nouns
    st.markdown(
        '<div class="card"><div class="card-title">고유명사 / 번역하지 않을 단어</div>'
        '<div class="card-sub">영어로도 그대로 쓸 브랜드명, 인명, 제품명 등을 쉼표로 구분해 입력하세요.</div>',
        unsafe_allow_html=True,
    )
    glossary = st.text_input(
        "glossary", placeholder="예: ChatGPT, POSCO, K-Beauty",
        value=st.session_state.glossary, label_visibility="collapsed",
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # Footer nav
    file_ready = bool(drive_url.strip() and _gdrive_file_id(drive_url.strip()))
    _, col_next = st.columns([1, 1])
    with col_next:
        go = st.button("다음 단계: 분류 →", type="primary", disabled=not file_ready,
                       use_container_width=True, key="btn_go")

    if go and file_ready:
        fid = _gdrive_file_id(drive_url.strip())
        with st.spinner("Google Drive에서 파일 다운로드 중..."):
            try:
                file_bytes = _download_gdrive(fid, is_slides=_gdrive_is_slides(drive_url.strip()))
            except Exception as e:
                st.error(f"다운로드 실패: {e}")
                st.stop()

        with st.spinner("텍스트 파싱 중..."):
            text_units = extract_text_units(file_bytes)
        if not text_units:
            st.error("번역 가능한 텍스트를 찾지 못했습니다. 파일을 확인해 주세요.")
            st.stop()

        with st.spinner("슬라이드 이미지 렌더링 중... (LibreOffice 필요)"):
            slide_imgs = render_slides_to_images(file_bytes)

        st.session_state.brand_name_ko = brand_ko
        st.session_state.brand_name_en = brand_en
        st.session_state.key_phrases = edited_kp.to_dict("records")
        st.session_state.glossary = glossary

        if ref_drive_url.strip():
            rfid = _gdrive_file_id(ref_drive_url.strip())
            if rfid:
                with st.spinner("참고 번역본 다운로드 중..."):
                    try:
                        ref_bytes = _download_gdrive(rfid)
                        st.session_state.ref_pptx_texts = extract_reference_texts(ref_bytes)
                    except Exception:
                        st.session_state.ref_pptx_texts = []
            else:
                st.session_state.ref_pptx_texts = []
        else:
            st.session_state.ref_pptx_texts = []

        if font_file is not None:
            try:
                from fontTools import ttLib
                tt = ttLib.TTFont(io.BytesIO(font_file.read()))
                fname = ""
                for record in tt["name"].names:
                    if record.nameID == 1:
                        try:
                            fname = record.toUnicode(); break
                        except Exception:
                            pass
                st.session_state.font_name = fname
            except Exception:
                st.session_state.font_name = ""
        else:
            st.session_state.font_name = ""

        st.session_state.file_bytes = file_bytes
        st.session_state.file_name = f"gdrive_{fid[:8]}_EN.pptx"
        from pptx import Presentation as _Prs
        st.session_state.slide_count = len(_Prs(io.BytesIO(file_bytes)).slides)
        st.session_state.text_units = text_units
        st.session_state.slide_images = slide_imgs
        st.session_state.classification_done = False
        st.session_state.active_classify_slide = 0
        st.session_state.stage = "classify"
        st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)


# ── Stage: classify ───────────────────────────────────────────────────────────
elif st.session_state.stage == "classify":

    if not st.session_state.classification_done:
        with st.spinner("AI가 텍스트를 분류하는 중..."):
            try:
                classified = classify_text_units(st.session_state.text_units, _build_context())
                st.session_state.classified_units = classified
                st.session_state.classification_done = True
            except Exception as e:
                st.error(f"분류 오류: {e}")
                st.stop()
        st.rerun()

    units = st.session_state.classified_units

    # Compute slide count — prefer the authoritative PPTX slide count
    slide_imgs = st.session_state.get("slide_images") or []
    n_from_pptx = st.session_state.get("slide_count") or 0
    n_from_imgs = len(slide_imgs)
    n_from_units = (max(u["slide_idx"] for u in units) + 1) if units else 0
    n_slides = max(n_from_pptx, n_from_imgs, n_from_units)
    active_slide = st.session_state.active_classify_slide
    if active_slide >= n_slides:
        active_slide = 0
        st.session_state.active_classify_slide = 0

    # Group units by slide
    slide_groups = defaultdict(list)
    for i, u in enumerate(units):
        slide_groups[u["slide_idx"]].append((i, u))

    # Page header
    st.markdown(
        '<div style="padding:28px 32px 16px;">'
        '<div class="page-title">슬라이드 분류</div>'
        '<div class="page-sub">감지된 텍스트 영역이 발표용 멘트인지 카피인지 확인하고, 태그를 클릭해 분류를 수정하세요.</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    rerun_needed = False
    col_left, col_center, col_right = st.columns([1.4, 4, 2])

    # Left: thumbnail strip
    with col_left:
        for s_idx in range(n_slides):
            is_active = (s_idx == active_slide)
            st.markdown(_thumb_html(s_idx, is_active), unsafe_allow_html=True)
            if st.button(f"슬라이드 {s_idx+1}", key=f"thumb_btn_{s_idx}",
                         use_container_width=True,
                         type="primary" if is_active else "secondary"):
                st.session_state.active_classify_slide = s_idx
                st.rerun()

    # Center: slide preview in bezel
    with col_center:
        st.markdown('<div class="bezel">', unsafe_allow_html=True)
        st.markdown(_slide_img_html(active_slide), unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # Right: legend + current slide text units + next button
    with col_right:
        st.markdown(
            '<div class="card">'
            '<div class="card-title" style="margin-bottom:14px;">분류 범례</div>'
            '<div class="legend-row"><div class="dot-y"></div>발표용 — 구두 발표 멘트</div>'
            '<div class="legend-row"><div class="dot-b"></div>카피 — 광고카피</div>'
            '</div>',
            unsafe_allow_html=True,
        )

        st.markdown(
            '<div class="card"><div class="card-title" style="margin-bottom:10px;">'
            f'슬라이드 {active_slide + 1}</div>',
            unsafe_allow_html=True,
        )

        slide_items = slide_groups.get(active_slide, [])
        if not slide_items:
            st.markdown('<div style="font-size:13px;color:#667085;">이 슬라이드에 번역할 텍스트가 없습니다.</div>',
                        unsafe_allow_html=True)
        else:
            # Bulk toggle row
            cb_all_p, cb_all_c = st.columns(2)
            with cb_all_p:
                if st.button("전체 발표용", key=f"all_pres_{active_slide}", use_container_width=True):
                    for i, u in slide_items:
                        st.session_state.classified_units[i]["category"] = "presentation"
                    st.rerun()
            with cb_all_c:
                if st.button("전체 카피", key=f"all_copy_{active_slide}", type="primary", use_container_width=True):
                    for i, u in slide_items:
                        st.session_state.classified_units[i]["category"] = "copy"
                    st.rerun()

            for i, u in slide_items:
                cat = u["category"]
                short = u["ko_text"][:45] + ("…" if len(u["ko_text"]) > 45 else "")
                c_tag, c_text = st.columns([1, 3])
                with c_tag:
                    tag_label = "발표용" if cat == "presentation" else "카피"
                    tag_type = "secondary" if cat == "presentation" else "primary"
                    if st.button(tag_label, key=f"tog_{u['id']}", type=tag_type, use_container_width=True):
                        st.session_state.classified_units[i]["category"] = (
                            "presentation" if cat == "copy" else "copy"
                        )
                        rerun_needed = True
                with c_text:
                    st.markdown(
                        f"<p style='font-size:12.5px;color:#101828;margin:0;"
                        f"padding:5px 0;line-height:1.4;'>{_html.escape(short)}</p>",
                        unsafe_allow_html=True,
                    )

        st.markdown('</div>', unsafe_allow_html=True)

        if rerun_needed:
            st.rerun()

        c_back, c_next = st.columns(2)
        with c_back:
            if st.button("← 이전", key="cl_back", use_container_width=True):
                st.session_state.stage = "upload"
                st.rerun()
        with c_next:
            if st.button("발표용 감수 →", key="cl_next", type="primary", use_container_width=True):
                pres = [u for u in st.session_state.classified_units if u["category"] == "presentation"]
                copy = [u for u in st.session_state.classified_units if u["category"] == "copy"]
                st.session_state.presentation_units = pres
                st.session_state.copy_units = copy
                st.session_state.translations_loaded = False
                st.session_state.copy_options_loaded = False
                st.session_state.current_pres_idx = 0
                st.session_state.current_copy_idx = 0
                st.session_state.copy_selections = {}
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

    if not pres_units:
        st.info("발표용 텍스트가 없습니다.")
        if st.button("카피 선택으로 →"):
            st.session_state.stage = "review_2b"
            st.rerun()
        st.stop()

    trans = st.session_state.presentation_translations

    # Group by slide and navigate slide-by-slide
    pres_by_slide = defaultdict(list)
    for u in pres_units:
        pres_by_slide[u["slide_idx"]].append(u)
    slide_keys = sorted(pres_by_slide.keys())
    total_slides = len(slide_keys)

    slide_pos = st.session_state.current_pres_idx
    if slide_pos >= total_slides:
        slide_pos = total_slides - 1
        st.session_state.current_pres_idx = slide_pos
    slide_idx = slide_keys[slide_pos]
    slide_units = pres_by_slide[slide_idx]

    # ── Two-column layout ───────────────────────────────────────────────────
    col_img, col_panel = st.columns([2, 3])

    with col_img:
        st.markdown(
            '<div class="sticky-slide" style="padding:16px 0 16px 24px;">'
            '<div class="bezel">',
            unsafe_allow_html=True,
        )
        st.markdown(_slide_img_html(slide_idx), unsafe_allow_html=True)
        st.markdown(
            f'<div class="bezel-caption">슬라이드 {slide_idx + 1}</div>'
            '</div></div>',
            unsafe_allow_html=True,
        )

    with col_panel:
        st.markdown('<div class="scroll-panel" style="padding:16px 8px 0 8px;">', unsafe_allow_html=True)
        for unit in slide_units:
            item = trans.get(unit["id"], {})
            en_text = item.get("en_text", "")
            notes = item.get("notes", "") or item.get("clarification", "")
            uid = unit["id"]

            # Korean source
            st.markdown(
                f'<div style="background:#F2F4F7;border-radius:8px 8px 0 0;'
                f'padding:10px 14px;font-size:13.5px;color:#344054;line-height:1.55;'
                f'border:1px solid #E4E7EC;border-bottom:none;">'
                f'{_html.escape(unit["ko_text"])}</div>',
                unsafe_allow_html=True,
            )
            # Editable English translation
            new_val = st.text_area(
                label="",
                value=en_text,
                key=f"edit_en_{uid}",
                height=80,
                label_visibility="collapsed",
            )
            if new_val != en_text:
                st.session_state.presentation_translations[uid] = {
                    **item, "en_text": new_val,
                }
            # Notes
            if notes:
                st.markdown(
                    f'<div style="background:#FFFBEA;border-radius:0 0 8px 8px;'
                    f'padding:8px 14px;font-size:12px;color:#667085;line-height:1.5;'
                    f'border:1px solid #E4E7EC;border-top:none;margin-bottom:14px;">'
                    f'📝 {_html.escape(notes)}</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown('<div style="margin-bottom:14px;"></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # ── AI chat refinement ──────────────────────────────────────────────────
    user_msg = st.chat_input(
        f"이 슬라이드 번역 수정 요청 ({slide_pos+1}/{total_slides} 슬라이드)", key="chat_2a"
    )
    if user_msg:
        with st.spinner("번역 수정 중..."):
            for unit in slide_units:
                item = trans.get(unit["id"], {})
                en_text = item.get("en_text", "")
                try:
                    refined = chat_refine_copy(unit["ko_text"], en_text, user_msg, _build_context())
                    st.session_state.presentation_translations[unit["id"]] = {
                        **item, "en_text": refined,
                    }
                except Exception:
                    pass
        st.rerun()

    # ── Navigation bar ──────────────────────────────────────────────────────
    c_prev, c_counter, c_next_btn, _, c_next_stage = st.columns([1, 1.5, 1, 3, 2])
    with c_prev:
        if st.button("＜", key="2a_prev", disabled=(slide_pos == 0), use_container_width=True):
            st.session_state.current_pres_idx -= 1
            st.rerun()
    with c_next_btn:
        if st.button("＞", key="2a_next", disabled=(slide_pos >= total_slides - 1), use_container_width=True):
            st.session_state.current_pres_idx += 1
            st.rerun()
    with c_counter:
        st.markdown(
            f'<div style="display:flex;align-items:center;height:38px;font-size:14px;'
            f'color:#667085;font-weight:500;">{slide_pos+1}/{total_slides} 슬라이드</div>',
            unsafe_allow_html=True,
        )
    with c_next_stage:
        next_label = "다음 단계: 카피 선택 →" if st.session_state.copy_units else "다음 단계: 다운로드 →"
        if st.button(next_label, key="2a_done", type="primary", use_container_width=True):
            st.session_state.stage = "review_2b" if st.session_state.copy_units else "download"
            st.rerun()


# ── Stage: review_2b ──────────────────────────────────────────────────────────
elif st.session_state.stage == "review_2b":

    copy_units = st.session_state.copy_units

    if not st.session_state.copy_options_loaded:
        with st.spinner("AI가 카피 옵션을 생성하는 중..."):
            try:
                opts = generate_copy_options(copy_units, _build_context())
                st.session_state.copy_options = opts
                st.session_state.copy_selections = {
                    u["id"]: opts.get(u["id"], {}).get("options", [""])[0]
                    for u in copy_units
                }
                st.session_state.copy_options_loaded = True
            except Exception as e:
                st.error(f"카피 생성 오류: {e}")
                st.stop()
        st.rerun()

    total = len(copy_units)
    idx = st.session_state.current_copy_idx
    if idx >= total:
        idx = total - 1
        st.session_state.current_copy_idx = idx
    if total == 0:
        st.info("카피 텍스트가 없습니다.")
        if st.button("다운로드 →"):
            st.session_state.stage = "download"
            st.rerun()
        st.stop()

    unit = copy_units[idx]
    unit_data = st.session_state.copy_options.get(unit["id"], {})
    options = unit_data.get("options", ["", "", ""])
    notes = unit_data.get("notes", "")
    recommendation = unit_data.get("recommendation", "")
    cultural_flag = unit_data.get("cultural_flag", "")
    clarification = unit_data.get("clarification", "")
    current_sel = st.session_state.copy_selections.get(unit["id"], options[0] if options else "")
    slide_idx = unit["slide_idx"]

    OPTS_META = [
        ("의역", "Feel, rhythm, impact 우선"),
        ("균형", "의미 + 자연스러운 영어"),
        ("직역", "원문 의미 중심"),
    ]

    # ── Two-column layout ───────────────────────────────────────────────────
    col_img, col_panel = st.columns([2, 3])

    with col_img:
        st.markdown(
            '<div style="padding:16px 0 16px 24px;">'
            '<div class="bezel">',
            unsafe_allow_html=True,
        )
        st.markdown(_slide_img_html(slide_idx), unsafe_allow_html=True)
        st.markdown(
            f'<div class="bezel-caption">슬라이드 {slide_idx + 1}</div>'
            '</div></div>',
            unsafe_allow_html=True,
        )

    with col_panel:
        st.markdown('<div style="padding:16px 24px 0 8px;">', unsafe_allow_html=True)

        # Korean source text header
        st.markdown(
            f'<div style="background:#F2F4F7;border-radius:8px;padding:10px 14px;'
            f'font-size:14px;color:#344054;line-height:1.55;margin-bottom:12px;">'
            f'{_html.escape(unit["ko_text"])}</div>',
            unsafe_allow_html=True,
        )

        # Notes
        note_content = notes or clarification or cultural_flag
        if note_content:
            st.markdown(
                f'<div class="notebox"><div class="notebox-icon">📝</div>'
                f'<div style="font-size:13px;">{_html.escape(note_content)}</div></div>',
                unsafe_allow_html=True,
            )

        # Copy option rows
        rerun_sel = False
        for i, (opt_name, opt_sub) in enumerate(OPTS_META):
            opt_text = options[i] if i < len(options) else ""
            is_sel = (current_sel == opt_text)
            sel_cls = "copyrow sel" if is_sel else "copyrow"
            star = "⭐" if is_sel else "☆"

            c_row, c_star = st.columns([10, 1])
            with c_row:
                st.markdown(
                    f'<div class="{sel_cls}">'
                    f'<div class="cr-label"><div class="cr-lname">{opt_name}</div>'
                    f'<div class="cr-lsub">{opt_sub}</div></div>'
                    f'<div class="cr-text">{_html.escape(opt_text)}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )
            with c_star:
                if st.button(star, key=f"star_{unit['id']}_{i}"):
                    st.session_state.copy_selections[unit["id"]] = opt_text
                    rerun_sel = True

        if rerun_sel:
            st.rerun()

        # Recommendation box
        if recommendation:
            st.markdown(
                f'<div class="recbox"><div style="font-size:18px;flex-shrink:0;">📌</div>'
                f'<div><strong>추천 이유:</strong> {_html.escape(recommendation)}</div></div>',
                unsafe_allow_html=True,
            )

        st.markdown('</div>', unsafe_allow_html=True)

    # ── AI chat ─────────────────────────────────────────────────────────────
    user_msg = st.chat_input(f"AI에게 카피 수정 요청하기 ({idx+1}/{total})", key="chat_2b")
    if user_msg:
        with st.spinner("카피 수정 중..."):
            try:
                refined = chat_refine_copy(unit["ko_text"], current_sel, user_msg, _build_context())
                st.session_state.copy_selections[unit["id"]] = refined
            except Exception as e:
                st.error(f"오류: {e}")
        st.rerun()

    # ── Navigation ───────────────────────────────────────────────────────────
    c_prev, c_counter, c_next_btn, _, c_next_stage = st.columns([1, 1.5, 1, 3, 2])
    with c_prev:
        if st.button("＜", key="2b_prev", disabled=(idx == 0), use_container_width=True):
            st.session_state.current_copy_idx -= 1
            st.rerun()
    with c_next_btn:
        if st.button("＞", key="2b_next", disabled=(idx >= total - 1), use_container_width=True):
            st.session_state.current_copy_idx += 1
            st.rerun()
    with c_counter:
        st.markdown(
            f'<div style="display:flex;align-items:center;height:38px;font-size:14px;'
            f'color:#667085;font-weight:500;">{idx+1}/{total} 카피</div>',
            unsafe_allow_html=True,
        )
    with c_next_stage:
        if st.button("다음 단계: 다운로드 →", key="2b_done", type="primary", use_container_width=True):
            st.session_state.stage = "download"
            st.rerun()


# ── Stage: download ───────────────────────────────────────────────────────────
elif st.session_state.stage == "download":

    if st.session_state.output_bytes is None:
        with st.spinner("번역본을 PPTX에 적용하는 중..."):
            pres_trans = {
                uid: v.get("en_text", "")
                for uid, v in st.session_state.presentation_translations.items()
            }
            all_translations = {**pres_trans, **st.session_state.copy_selections}
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

    st.markdown('<div class="page">', unsafe_allow_html=True)
    st.markdown('<div class="page-title">번역 완료</div>', unsafe_allow_html=True)
    n_pres = len(st.session_state.presentation_units)
    n_copy = len(st.session_state.copy_units)
    st.markdown(
        f'<div class="page-sub">발표용 텍스트 {n_pres}개 번역 · 광고 카피 {n_copy}개 트랜스크리에이션 완료</div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.success("영문 번역본이 준비되었습니다!")
    st.download_button(
        label="영문 PPT 다운로드",
        data=st.session_state.output_bytes,
        file_name=st.session_state.file_name,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
    )
    st.markdown('</div>', unsafe_allow_html=True)

    if st.button("처음으로 돌아가기"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
