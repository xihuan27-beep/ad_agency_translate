import streamlit as st
import pandas as pd
import io
import re
import html as _html
import requests
import base64
from collections import defaultdict

from pptx_utils import extract_text_units, apply_translations, extract_reference_texts, render_slides_to_images
from docx_utils import extract_text_from_docx, apply_translations_to_docx
from ai_utils import (
    classify_text_units,
    translate_presentation_texts,
    generate_copy_options,
    chat_modify_presentation,
    chat_refine_copy,
    check_copy_grammar,
    translate_en_to_ko,
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
  --cp:  #0C2790;   /* navy primary */
  --cpa: #0f31ad;   /* navy hover */
  --cbg: #F5F7FC;   /* cool off-white page bg */
  --ct:  #0D1B2A;   /* primary text */
  --cm:  #4A5B7A;   /* secondary text */
  --cb:  #DCE3F0;   /* border */
  --cw:  #FFFFFF;
  --cbl: #E8ECFA;   /* light navy tint */
  --copy-accent: #2ECC46;  /* copy tag accent (green) */
  --copy-bg:     #EAFBEE;  /* copy tag bg */
  --copy-border: #A9E8B5;  /* copy tag border */
  --cs:  0 1px 3px rgba(12,39,144,0.06), 0 4px 12px rgba(12,39,144,0.05);  /* card shadow */
  --r:   10px;
  --font: 'Pretendard Variable', -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif;
}
body, .stApp, .main { background: var(--cbg) !important; font-family: var(--font) !important; }
section[data-testid="stMain"] { background: var(--cbg) !important; }
* { box-sizing: border-box; }

/* Data editor */
[data-testid="stDataFrameResizable"] {
  border: 1px solid var(--cb) !important; border-radius: 8px !important;
  overflow: hidden !important; background: var(--cw) !important;
}
.glideDataEditor, .dvn-scroller { background: var(--cw) !important; }

/* File uploader */
[data-testid="stFileUploadDropzone"] {
  background: var(--cbg) !important; border: 1.5px dashed var(--cb) !important;
  border-radius: 8px !important;
}
[data-testid="stFileUploadDropzone"] section { background: transparent !important; }
[data-testid="stFileUploadDropzone"] p,
[data-testid="stFileUploadDropzone"] span { color: var(--cm) !important; }

/* Container transparency */
[data-testid="stVerticalBlock"] > [data-testid="stVerticalBlock"] { background: transparent !important; }
[data-testid="column"] { background: transparent !important; }

/* ── Top bar ── */
.topbar {
  display: flex; align-items: center; gap: 12px;
  padding: 0 32px; height: 58px;
  background: var(--cw); border-bottom: 1px solid var(--cb);
}
.topbar-logo {
  width: 32px; height: 32px; border-radius: 8px; background: var(--cp);
  display: flex; align-items: center; justify-content: center; flex-shrink: 0;
}
.topbar-title { font-size: 14.5px; font-weight: 600; color: var(--ct); letter-spacing: -0.01em; }

/* ── Step rail ── */
.steprail {
  display: flex; align-items: stretch; border-bottom: 1px solid var(--cb);
  background: var(--cw); padding: 0 28px; height: 48px;
}
.step {
  display: flex; align-items: center; gap: 7px; padding: 0 16px;
  font-size: 13px; font-weight: 500; color: var(--cm);
  position: relative; white-space: nowrap;
}
.step.active { color: var(--ct); font-weight: 600; }
.step.active::after {
  content: ''; position: absolute; bottom: 0; left: 0; right: 0;
  height: 2px; background: var(--cp); border-radius: 2px 2px 0 0;
}
.step-n {
  width: 22px; height: 22px; border-radius: 50%;
  display: flex; align-items: center; justify-content: center;
  font-size: 11px; font-weight: 700; flex-shrink: 0;
  border: 1.5px solid var(--cb); background: transparent; color: var(--cm);
}
.step.active .step-n { background: var(--cp); color: #fff; border-color: var(--cp); }
.step.done .step-n   { background: var(--cbl); color: var(--cpa); border-color: #C7D2F0; }
.step.done { color: var(--cm); }

/* ── Card ── */
/* HTML-only card (no interactive widgets inside) */
.card {
  background: var(--cw); border-radius: var(--r);
  box-shadow: var(--cs); padding: 22px 24px; margin-bottom: 14px;
}
/* st.container(border=True, key="card_...") → restyle as card (used when widgets live inside).
   Streamlit no longer exposes a stable "border wrapper" testid, so every card container
   is given an explicit key and targeted via its st-key-* class instead. */
[class*="st-key-card_"] {
  background: var(--cw) !important;
  border: none !important;
  border-radius: var(--r) !important;
  box-shadow: var(--cs) !important;
  padding: 18px 22px !important;
  margin-bottom: 14px !important;
}
.card-title { font-size: 14px; font-weight: 600; color: var(--ct); margin-bottom: 8px; }
.card-sub   { font-size: 12.5px; color: var(--cm); margin-bottom: 14px; line-height: 1.55; }
.field-label { font-size: 12px; color: var(--cm); font-weight: 500; margin-bottom: 5px; }

/* ── Page header ── */
.page { padding: 28px 32px 40px; }
.page-title { font-size: 20px; font-weight: 700; color: var(--ct); margin-bottom: 4px; letter-spacing: -0.02em; }
.page-sub   { font-size: 13.5px; color: var(--cm); margin-bottom: 20px; }

/* ── Slide bezel ── */
.bezel {
  background: #F0EFEC; border-radius: 12px; padding: 16px; margin-bottom: 0;
  border: 1px solid var(--cb);
}
.bezel img { border-radius: 6px; width: 100%; display: block; box-shadow: 0 2px 8px rgba(0,0,0,0.12); }
.bezel-placeholder {
  border-radius: 6px; background: #E8E6E2; width: 100%; aspect-ratio: 16/9;
  display: flex; align-items: center; justify-content: center;
  color: #9CA3AF; font-size: 14px;
}
.bezel-caption { color: var(--cm); font-size: 12.5px; margin-top: 10px; text-align: center; }

/* ── Notebox ── */
.notebox {
  background: #FAFAF8; border: 1px solid var(--cb); border-radius: 8px;
  padding: 12px 14px; display: flex; gap: 10px; align-items: flex-start;
  font-size: 13px; line-height: 1.6; color: var(--cm); margin: 10px 0;
}
.notebox-icon { font-size: 16px; flex-shrink: 0; margin-top: 1px; }
.transbox {
  background: var(--cbl); border-radius: 8px; padding: 20px 24px;
  text-align: center; font-size: 20px; font-weight: 700;
  color: var(--cpa); line-height: 1.4; margin: 12px 0;
}

/* ── Copy option rows ── */
.copyrow {
  display: flex; flex-direction: column; gap: 6px;
  padding: 13px 15px; border-radius: 8px; border: 1px solid var(--cb);
  margin-bottom: 8px; background: var(--cw); box-shadow: var(--cs);
  transition: background .12s, border-color .12s;
}
.copyrow.sel { background: var(--cp); border-color: var(--cp); }
.cr-label { display: flex; align-items: center; gap: 8px; }
.cr-lname {
  font-size: 10px; font-weight: 700; letter-spacing: .04em; text-transform: uppercase;
  padding: 2px 8px; border-radius: 12px; background: var(--cbg); color: var(--cm);
  border: 1px solid var(--cb);
}
.copyrow.sel .cr-lname { background: rgba(255,255,255,.14); color: rgba(255,255,255,.85); border-color: rgba(255,255,255,.25); }
.cr-lsub  { font-size: 11px; color: var(--cm); }
.copyrow.sel .cr-lsub { color: rgba(255,255,255,.65); }
.cr-text  { font-size: 14.5px; font-weight: 500; color: var(--ct); line-height: 1.55; }
.copyrow.sel .cr-text { color: #fff; }

/* ── Recommendation ── */
.recbox {
  background: var(--cbl); border-radius: 8px; padding: 12px 16px;
  display: flex; gap: 10px; font-size: 13px; line-height: 1.65;
  color: var(--ct); margin-bottom: 12px;
}

/* ── Tags ── */
.tag-p { display:inline-flex;align-items:center;gap:4px;padding:2px 9px;
  background:var(--cbl);color:var(--cp);border-radius:20px;font-size:11.5px;font-weight:600; }
.tag-c { display:inline-flex;align-items:center;gap:4px;padding:2px 9px;
  background:var(--copy-bg);color:var(--copy-accent);border-radius:20px;font-size:11.5px;font-weight:600; }

/* ── Legend ── */
.legend-row { display:flex;align-items:center;gap:8px;margin-bottom:8px;font-size:13px;color:var(--ct); }
.dot-y { width:10px;height:10px;border-radius:50%;background:var(--cp);flex-shrink:0; }
.dot-b { width:10px;height:10px;border-radius:50%;background:var(--copy-accent);flex-shrink:0; }

/* ── Classify item ── */
.cl-item {
  display:flex;align-items:flex-start;justify-content:space-between;gap:10px;
  padding:9px 12px;border-radius:7px;background:var(--cbg);
  border:1px solid var(--cb);margin-bottom:7px;
}
.cl-item-text { font-size:13px;color:var(--ct);line-height:1.5;flex:1; }

/* ── Buttons ── */
div.stButton > button {
  border-radius: 7px !important; font-weight: 500 !important;
  font-size: 13.5px !important;
}
div.stButton > button[kind="primary"] {
  background: var(--cp) !important; border-color: var(--cp) !important; color: #fff !important;
}
div.stButton > button[kind="primary"]:hover {
  background: var(--cpa) !important; border-color: var(--cpa) !important;
}

/* ── Input fields ── */
.stTextInput > div > div > input,
.stTextArea > div > div > textarea {
  background: #fff !important; color: var(--ct) !important;
  border: 1px solid var(--cb) !important; border-radius: 7px !important;
  font-size: 14px !important; box-shadow: none !important;
}
.stTextInput > div > div > input::placeholder,
.stTextArea > div > div > textarea::placeholder { color: #9CA3AF !important; }
.stTextInput > div > div > input:focus,
.stTextArea > div > div > textarea:focus {
  border-color: var(--cp) !important;
  box-shadow: 0 0 0 3px rgba(12,39,144,0.15) !important;
}
.stSelectbox > div > div > div {
  background: #fff !important; color: var(--ct) !important;
  border: 1px solid var(--cb) !important; border-radius: 7px !important;
}
[data-baseweb="input"]           { background: #fff !important; }
[data-baseweb="input"] input     { background: transparent !important; color: var(--ct) !important; }
[data-baseweb="textarea"]        { background: #fff !important; }
[data-baseweb="textarea"] textarea { background: transparent !important; color: var(--ct) !important; }
.stTextInput label, .stTextArea label, .stSelectbox label {
  color: var(--cm) !important; font-size: 12px !important; font-weight: 500 !important;
}

/* ── Misc ── */
.navrow { display:flex;align-items:center;justify-content:space-between;padding:14px 0;gap:10px; }
[data-testid="stHorizontalBlock"] { align-items: flex-start !important; }
.review-img-panel { padding: 16px 0 16px 20px; position: sticky; top: 70px; }

/* Sticky slide image panel — works in classify center col and review left col */
[data-testid="column"]:has(.bezel) {
  position: sticky !important;
  top: 106px !important;   /* topbar 58px + steprail 48px */
  align-self: flex-start !important;
}

/* ── Thumbnail strip: minimal "선택" pill under each numbered thumbnail ── */
.st-key-thumb_strip div.stButton { margin: -6px 0 12px 21px; }
.st-key-thumb_strip div.stButton > button {
  font-size: 10.5px !important; padding: 1px 0 !important; height: 20px !important;
  min-height: 20px !important; border-color: transparent !important;
  color: var(--cm) !important; background: transparent !important; box-shadow: none !important;
}
.st-key-thumb_strip div.stButton > button:hover { color: var(--cp) !important; }
.st-key-thumb_strip div.stButton > button[kind="primary"] {
  background: var(--cbl) !important; color: var(--cp) !important; border-color: transparent !important;
}

/* ── Classify item cards: same card look, tighter padding than a full section card ── */
[class*="st-key-item_card_"] {
  background: var(--cw) !important;
  border: none !important;
  border-radius: 8px !important;
  box-shadow: var(--cs) !important;
  padding: 11px 13px !important;
  margin-bottom: 8px !important;
}

/* ── Direction picker: two large clickable cards (flags) ── */
.dir-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
.dir-card {
  padding: 22px 16px; border-radius: 10px; border: 2px solid var(--cb);
  background: var(--cw); text-align: center; transition: border-color .15s, background .15s, color .15s;
}
.dir-card.selected { border-color: var(--cp); background: var(--cp); color: #fff; }
.dir-arrow { font-size: 22px; display: block; margin-bottom: 8px; }
.dir-title { font-size: 14px; font-weight: 600; }
.dir-sub { font-size: 11.5px; opacity: .7; margin-top: 3px; }
.dir-picked {
  text-align: center; font-size: 12px; font-weight: 600; color: var(--cp);
  padding: 7px 0 14px;
}
[class*="st-key-dir_pick_"] { margin: 6px 0 14px; }
[class*="st-key-dir_pick_"] button {
  font-size: 12.5px !important; color: var(--cm) !important;
  background: var(--cw) !important; border-color: var(--cb) !important;
}
[class*="st-key-dir_pick_"] button:hover { color: var(--cp) !important; border-color: var(--cp) !important; }

/* ── Classify tag buttons: 발표용 = navy tint, 카피 = green tint ── */
[class*="st-key-tog_pres_"] button, [class*="st-key-all_pres_"] button {
  background: var(--cbl) !important; color: var(--cp) !important; border-color: rgba(12,39,144,.2) !important;
}
[class*="st-key-tog_pres_"] button:hover, [class*="st-key-all_pres_"] button:hover { border-color: var(--cp) !important; }
[class*="st-key-tog_copy_"] button, [class*="st-key-all_copy_btn_"] button {
  background: var(--copy-bg) !important; color: var(--copy-accent) !important; border-color: var(--copy-border) !important;
}
[class*="st-key-tog_copy_"] button:hover, [class*="st-key-all_copy_btn_"] button:hover { border-color: var(--copy-accent) !important; }

/* ── Review KO/EN pair block: white KO on top, navy EN on bottom, one seamless card ── */
[class*="st-key-pair_"] {
  background: var(--cw) !important; border: none !important; border-radius: var(--r) !important;
  box-shadow: var(--cs) !important; padding: 0 !important; margin-bottom: 0 !important;
  overflow: hidden !important; gap: 0 !important;
}
[class*="st-key-pair_"] [data-testid="stElementContainer"],
[class*="st-key-pair_"] [data-testid="stTextArea"] { margin: 0 !important; }
[class*="st-key-pair_"] textarea {
  background: var(--cp) !important; color: #fff !important; border: none !important;
  border-radius: 0 !important; box-shadow: none !important; padding: 13px 15px !important;
  font-size: 13.5px !important; line-height: 1.7 !important; font-family: var(--font) !important;
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
        "direction": "ko_en",
        "file_type": "pptx",
        "en_ko_translations": {},
        "en_ko_loaded": False,
        "copy_grammar_results": {},
        "excluded_unit_ids": [],
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
        r"/document/d/([a-zA-Z0-9_-]+)",
        r"/spreadsheets/d/([a-zA-Z0-9_-]+)",
    ]
    for p in patterns:
        m = re.search(p, url)
        if m:
            return m.group(1)
    return None

def _gdrive_is_slides(url: str) -> bool:
    return bool(re.search(r"/presentation/d/", url))

def _gdrive_is_docs(url: str) -> bool:
    return bool(re.search(r"/document/d/", url))

def _is_zip(data: bytes) -> bool:
    """PPTX and DOCX are ZIP files — both start with PK magic bytes."""
    return len(data) >= 4 and data[:4] == b'PK\x03\x04'

def _gdrive_confirm_url(html_text: str, file_id: str) -> str | None:
    """Extract confirmed download URL from Google's virus-scan warning page."""
    m = re.search(r'confirm=([0-9A-Za-z_-]+)', html_text)
    if m:
        return f"https://drive.google.com/uc?export=download&id={file_id}&confirm={m.group(1)}"
    return None

def _download_gdrive(file_id: str, is_slides: bool = False, is_docs: bool = False) -> bytes:
    session = requests.Session()
    errors = []

    if is_slides:
        url = f"https://docs.google.com/presentation/d/{file_id}/export/pptx"
        try:
            resp = session.get(url, timeout=300)
            if resp.status_code == 200 and _is_zip(resp.content):
                return resp.content
            errors.append(f"Slides export: HTTP {resp.status_code}, size={len(resp.content)}")
        except Exception as e:
            errors.append(f"Slides export: {e}")

    if is_docs:
        url = f"https://docs.google.com/document/d/{file_id}/export?format=docx"
        try:
            resp = session.get(url, timeout=300)
            if resp.status_code == 200 and _is_zip(resp.content):
                return resp.content
            errors.append(f"Docs export: HTTP {resp.status_code}, size={len(resp.content)}")
        except Exception as e:
            errors.append(f"Docs export: {e}")

    # Direct usercontent download (works for regular Drive files)
    url = f"https://drive.usercontent.google.com/download?id={file_id}&export=download&authuser=0&confirm=t"
    try:
        resp = session.get(url, timeout=300)
        if resp.status_code == 200 and _is_zip(resp.content):
            return resp.content
        # Google may return an HTML virus-scan confirmation page
        if resp.status_code == 200 and resp.content[:1] == b'<':
            confirm = _gdrive_confirm_url(resp.text, file_id)
            if confirm:
                resp2 = session.get(confirm, timeout=300)
                if resp2.status_code == 200 and _is_zip(resp2.content):
                    return resp2.content
        errors.append(f"usercontent: HTTP {resp.status_code}, size={len(resp.content)}")
    except Exception as e:
        errors.append(f"usercontent: {e}")

    # Slides export fallback (try even for non-Slides URLs — works if it happens to be Slides)
    if not is_slides:
        url = f"https://docs.google.com/presentation/d/{file_id}/export/pptx"
        try:
            resp = session.get(url, timeout=300)
            if resp.status_code == 200 and _is_zip(resp.content):
                return resp.content
            errors.append(f"Slides fallback: HTTP {resp.status_code}")
        except Exception as e:
            errors.append(f"Slides fallback: {e}")

    # Legacy /uc endpoint with cookie + HTML confirm handling
    url = f"https://drive.google.com/uc?export=download&id={file_id}"
    try:
        resp = session.get(url, timeout=300)
        for key, value in resp.cookies.items():
            if key.startswith("download_warning"):
                url = f"https://drive.google.com/uc?export=download&id={file_id}&confirm={value}"
                resp = session.get(url, timeout=300)
                break
        if not _is_zip(resp.content) and resp.content[:1] == b'<':
            confirm = _gdrive_confirm_url(resp.text, file_id)
            if confirm:
                resp = session.get(confirm, timeout=300)
        if resp.status_code == 200 and _is_zip(resp.content):
            return resp.content
        errors.append(f"legacy uc: HTTP {resp.status_code}, size={len(resp.content)}")
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
    border = "2px solid #0C2790" if active else "2px solid #DCE3F0"
    inner = _slide_img_html(slide_idx)
    return (
        f'<div style="display:flex;align-items:flex-start;gap:7px;">'
        f'<span style="font-size:11px;color:#8A9BB8;font-weight:500;width:14px;'
        f'text-align:right;padding-top:5px;flex-shrink:0;">{slide_idx+1}</span>'
        f'<div style="flex:1;min-width:0;border:{border};border-radius:6px;overflow:hidden;">'
        f'{inner}'
        f'</div>'
        f'</div>'
    )


# ── Chrome: top bar + step rail ───────────────────────────────────────────────
STAGES_KO_EN = ["upload", "classify", "review_2a", "review_2b", "download"]
LABELS_KO_EN = ["업로드", "분류", "발표용 감수", "카피 선택", "다운로드"]
STAGES_EN_KO = ["upload", "en_ko", "download"]
LABELS_EN_KO = ["업로드", "번역", "다운로드"]

def _current_stages() -> tuple[list, list]:
    if st.session_state.direction == "en_ko":
        return STAGES_EN_KO, LABELS_EN_KO
    return STAGES_KO_EN, LABELS_KO_EN

def _render_chrome():
    stages, labels = _current_stages()
    curr_stage = st.session_state.stage
    curr_idx = stages.index(curr_stage) if curr_stage in stages else 0
    st.markdown(
        '<div class="topbar">'
        '<div class="topbar-logo">'
        '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">'
        '<rect x="2" y="3" width="20" height="15" rx="2" stroke="white" stroke-width="2" fill="none"/>'
        '<path d="M8 22h8M12 18v4" stroke="white" stroke-width="2" stroke-linecap="round"/>'
        '</svg></div>'
        '<span class="topbar-title">Agency Deck Translator</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    items = ""
    for i, label in enumerate(labels):
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

    # Direction selector — two large clickable cards (matches approved mockup)
    st.markdown('<div class="section-eyebrow">번역 방향</div>', unsafe_allow_html=True)
    _dir_col1, _dir_col2 = st.columns(2)
    with _dir_col1:
        _sel = st.session_state.direction == "ko_en"
        st.markdown(
            f'<div class="dir-card{" selected" if _sel else ""}">'
            f'<span class="dir-arrow">🇰🇷 → 🇺🇸</span>'
            f'<div class="dir-title">한국어 → 영어</div>'
            f'<div class="dir-sub">Korean to English</div></div>',
            unsafe_allow_html=True,
        )
        if _sel:
            st.markdown('<div class="dir-picked">✓ 선택됨</div>', unsafe_allow_html=True)
        elif st.button("이 방향 선택", key="dir_pick_ko_en", use_container_width=True):
            st.session_state.direction = "ko_en"
            st.rerun()
    with _dir_col2:
        _sel = st.session_state.direction == "en_ko"
        st.markdown(
            f'<div class="dir-card{" selected" if _sel else ""}">'
            f'<span class="dir-arrow">🇺🇸 → 🇰🇷</span>'
            f'<div class="dir-title">영어 → 한국어</div>'
            f'<div class="dir-sub">English to Korean</div></div>',
            unsafe_allow_html=True,
        )
        if _sel:
            st.markdown('<div class="dir-picked">✓ 선택됨</div>', unsafe_allow_html=True)
        elif st.button("이 방향 선택", key="dir_pick_en_ko", use_container_width=True):
            st.session_state.direction = "en_ko"
            st.rerun()

    _is_en_ko = st.session_state.direction == "en_ko"

    if _is_en_ko:
        st.markdown('<div class="page-title">영어 문서 업로드</div>', unsafe_allow_html=True)
        st.markdown('<div class="page-sub">번역할 영어 PPTX 또는 Word(.docx) 파일의 Google Drive 링크를 입력하세요.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="page-title">PPT 파일 업로드 및 캠페인 브리프</div>', unsafe_allow_html=True)
        st.markdown('<div class="page-sub">번역할 한국어 PPTX 파일과 캠페인 정보를 입력하세요.</div>', unsafe_allow_html=True)

    # Card 1: file
    _file_card_title = "영어 파일 (PPTX / Word)" if _is_en_ko else "한국어 PPTX 파일"
    _file_placeholder = "https://drive.google.com/file/d/...  또는  https://docs.google.com/document/d/..."
    with st.container(border=True, key="card_file"):
        st.markdown(f'<div class="card-title">{_file_card_title}</div>', unsafe_allow_html=True)
        st.markdown('<div class="field-label">☁ Google Drive / Google Docs 링크</div>', unsafe_allow_html=True)
        c1, c2 = st.columns([5, 1])
        with c1:
            drive_url = st.text_input(
                "drive_url", placeholder=_file_placeholder,
                label_visibility="collapsed", key="up_drive_url",
            )
        with c2:
            fetch_clicked = st.button("가져오기", key="btn_fetch_main", use_container_width=True, type="primary")
        if drive_url.strip() and not _gdrive_file_id(drive_url.strip()):
            st.warning("올바른 Google Drive / Google Docs 링크가 아닙니다.")

    # Card 2: Translation quality info (brand name + term mapping, grouped as one section)
    if _is_en_ko:
        _kp_title = "주요 용어 매핑 (영어 → 한국어)"
        _kp_sub = "자주 쓰이는 영어 표현의 선호 한국어 번역을 지정합니다."
        _kp_col1, _kp_col2 = "영어", "한국어"
        _kp_label1, _kp_label2 = "영어 표현", "한국어 번역 (선호)"
        _kp_init = st.session_state.key_phrases if st.session_state.key_phrases else [{"영어": "", "한국어": ""}]
    else:
        _kp_title = "주요 용어 매핑 (한국어 → 영어)"
        _kp_sub = "자주 쓰는 표현의 선호 번역을 지정합니다. 행 추가 버튼으로 한 줄을 추가하세요."
        _kp_col1, _kp_col2 = "한국어", "영어"
        _kp_label1, _kp_label2 = "한국어 표현", "영어 번역 (선호)"
        _kp_init = st.session_state.key_phrases if st.session_state.key_phrases else [{"한국어": "", "영어": ""}]
    with st.container(border=True, key="card_quality"):
        st.markdown(
            '<div class="card-title">번역 품질 향상 정보 <span style="font-weight:400;opacity:.7;">(선택)</span></div>',
            unsafe_allow_html=True,
        )
        st.markdown('<div class="field-label" style="margin-top:4px;">브랜드명</div>', unsafe_allow_html=True)
        c_ko, c_en = st.columns(2)
        with c_ko:
            brand_ko = st.text_input("brand_ko", placeholder="브랜드명 (한국어) — 예: 삼성전자",
                value=st.session_state.brand_name_ko, label_visibility="collapsed")
        with c_en:
            brand_en = st.text_input("brand_en", placeholder="Brand name (EN) — e.g. Samsung Electronics",
                value=st.session_state.brand_name_en, label_visibility="collapsed")

        st.markdown(f'<div class="field-label" style="margin-top:14px;">{_kp_title}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="card-sub">{_kp_sub}</div>', unsafe_allow_html=True)
        edited_kp = st.data_editor(
            pd.DataFrame(_kp_init),
            column_config={
                _kp_col1: st.column_config.TextColumn(_kp_label1, width="large"),
                _kp_col2: st.column_config.TextColumn(_kp_label2, width="large"),
            },
            num_rows="dynamic", hide_index=True, use_container_width=True, key="kp_editor",
        )

    # Cards only shown for ko→en
    ref_drive_url = ""
    font_file = None
    if not _is_en_ko:
        # Card 4: Reference PPTX (optional)
        with st.container(border=True, key="card_ref"):
            st.markdown(
                '<div class="card-title">이전 번역본 참고 (선택)</div>'
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

        # Card 5: Font (optional)
        with st.container(border=True, key="card_font"):
            st.markdown(
                '<div class="card-title">영어 폰트 (선택)</div>'
                '<div class="card-sub">번역된 텍스트에 적용할 TTF/OTF 폰트 파일을 업로드하세요.</div>',
                unsafe_allow_html=True,
            )
            font_file = st.file_uploader("폰트 파일 (선택, TTF/OTF)", type=["ttf", "otf"], key="font_uploader",
                                          label_visibility="collapsed")
            if font_file:
                st.caption(f"업로드: {font_file.name}")

    # Card 6: Proper nouns
    _noun_sub = "번역하지 않고 그대로 쓸 브랜드명, 인명, 제품명 등을 쉼표로 구분해 입력하세요."
    with st.container(border=True, key="card_glossary"):
        st.markdown(
            f'<div class="card-title">고유명사 / 번역하지 않을 단어</div>'
            f'<div class="card-sub">{_noun_sub}</div>',
            unsafe_allow_html=True,
        )
        glossary = st.text_input(
            "glossary", placeholder="예: ChatGPT, POSCO, K-Beauty",
            value=st.session_state.glossary, label_visibility="collapsed",
        )

    # Footer nav
    file_ready = bool(drive_url.strip() and _gdrive_file_id(drive_url.strip()))
    _next_label = "다음 단계: 번역 →" if _is_en_ko else "다음 단계: 분류 →"
    _, col_next = st.columns([1, 1])
    with col_next:
        go = st.button(_next_label, type="primary", disabled=not file_ready,
                       use_container_width=True, key="btn_go")

    if go and file_ready:
        _url = drive_url.strip()
        fid = _gdrive_file_id(_url)
        _is_slides = _gdrive_is_slides(_url)
        _is_docs = _gdrive_is_docs(_url)
        with st.spinner("Google Drive에서 파일 다운로드 중..."):
            try:
                file_bytes = _download_gdrive(fid, is_slides=_is_slides, is_docs=_is_docs)
            except Exception as e:
                st.error(f"다운로드 실패: {e}")
                st.stop()

        # Validate file is actually a ZIP (PPTX/DOCX are ZIP-based)
        if not _is_zip(file_bytes):
            st.error(
                "파일을 올바르게 다운로드하지 못했습니다. "
                "Google Drive 파일 공유 설정을 확인해 주세요:\n\n"
                "1. Google Drive에서 파일 우클릭 → 공유\n"
                "2. '링크가 있는 모든 사용자' 또는 '뷰어'로 설정\n"
                "3. 링크 복사 후 다시 시도"
            )
            st.stop()

        # Detect file type from URL or content header
        _file_type = "docx" if (_is_docs or _url.lower().endswith(".docx")) else "pptx"

        with st.spinner("텍스트 파싱 중..."):
            if _file_type == "docx":
                text_units = extract_text_from_docx(file_bytes)
            else:
                text_units = extract_text_units(file_bytes)
        if not text_units:
            st.error("번역 가능한 텍스트를 찾지 못했습니다. 파일을 확인해 주세요.")
            st.stop()

        # Slide images only for PPTX
        slide_imgs = []
        slide_count = 0
        if _file_type == "pptx":
            with st.spinner("슬라이드 이미지 렌더링 중... (LibreOffice 필요)"):
                slide_imgs = render_slides_to_images(file_bytes)
            from pptx import Presentation as _Prs
            slide_count = len(_Prs(io.BytesIO(file_bytes)).slides)

        st.session_state.brand_name_ko = brand_ko
        st.session_state.brand_name_en = brand_en
        st.session_state.key_phrases = edited_kp.to_dict("records")
        st.session_state.glossary = glossary
        st.session_state.file_type = _file_type

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
        _suffix = "KO" if _is_en_ko else "EN"
        _ext = _file_type
        st.session_state.file_name = f"gdrive_{fid[:8]}_{_suffix}.{_ext}"
        st.session_state.slide_count = slide_count
        st.session_state.text_units = text_units
        st.session_state.slide_images = slide_imgs
        st.session_state.output_bytes = None

        if _is_en_ko:
            st.session_state.en_ko_loaded = False
            st.session_state.en_ko_translations = {}
            st.session_state.stage = "en_ko"
        else:
            st.session_state.classification_done = False
            st.session_state.active_classify_slide = 0
            st.session_state.stage = "classify"
        st.rerun()


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

    # Left: thumbnail strip — independently scrollable
    with col_left:
        with st.container(height=650, key="thumb_strip"):
            for s_idx in range(n_slides):
                is_active = (s_idx == active_slide)
                st.markdown(_thumb_html(s_idx, is_active), unsafe_allow_html=True)
                if st.button("선택", key=f"thumb_btn_{s_idx}",
                             use_container_width=True,
                             type="primary" if is_active else "secondary"):
                    st.session_state.active_classify_slide = s_idx
                    st.rerun()

    # Center: slide preview in bezel
    with col_center:
        st.markdown('<div class="bezel">', unsafe_allow_html=True)
        st.markdown(_slide_img_html(active_slide), unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown(
            '<div style="font-size:11px;color:#9CA3AF;text-align:center;margin-top:6px;">'
            '미리보기는 서버 폰트 제한으로 실제 PPT와 다를 수 있습니다'
            '</div>',
            unsafe_allow_html=True,
        )

    # Right: legend + current slide text units (independently scrollable); nav buttons below
    with col_right:
        with st.container(height=650):
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
                    if st.button("전체 카피", key=f"all_copy_btn_{active_slide}", type="secondary", use_container_width=True):
                        for i, u in slide_items:
                            st.session_state.classified_units[i]["category"] = "copy"
                        st.rerun()

                excluded_ids = set(st.session_state.excluded_unit_ids)
                for pos, (i, u) in enumerate(slide_items):
                    cat = u["category"]
                    is_excluded = u["id"] in excluded_ids
                    short = u["ko_text"][:45] + ("…" if len(u["ko_text"]) > 45 else "")
                    has_next = pos + 1 < len(slide_items)

                    with st.container(border=True, key=f"item_card_{u['id']}"):
                        c_tag, c_text, c_merge, c_excl = st.columns([1, 3, 0.5, 0.5])
                        with c_tag:
                            if is_excluded:
                                st.markdown(
                                    '<p style="font-size:11px;color:#9CA3AF;margin:0;padding:5px 0;">제외됨</p>',
                                    unsafe_allow_html=True,
                                )
                            else:
                                tag_label = "발표용" if cat == "presentation" else "카피"
                                tag_key = f"tog_pres_{u['id']}" if cat == "presentation" else f"tog_copy_{u['id']}"
                                if st.button(tag_label, key=tag_key, type="secondary", use_container_width=True):
                                    st.session_state.classified_units[i]["category"] = (
                                        "presentation" if cat == "copy" else "copy"
                                    )
                                    rerun_needed = True
                        with c_text:
                            _tc = "#9CA3AF" if is_excluded else "#101828"
                            _td = "line-through" if is_excluded else "none"
                            st.markdown(
                                f"<p style='font-size:12.5px;color:{_tc};margin:0;"
                                f"padding:5px 0;line-height:1.4;text-decoration:{_td};'>"
                                f"{_html.escape(short)}</p>",
                                unsafe_allow_html=True,
                            )
                        with c_merge:
                            if has_next and not is_excluded:
                                if st.button("↕", key=f"merge_{u['id']}", use_container_width=True,
                                             help="다음 항목과 합치기"):
                                    j, _ = slide_items[pos + 1]
                                    cu = st.session_state.classified_units[i]
                                    cv = st.session_state.classified_units[j]
                                    merged = cu["ko_text"] + " " + cv["ko_text"]
                                    st.session_state.classified_units[i]["ko_text"] = merged
                                    st.session_state.classified_units[i]["shape_text"] = merged
                                    del st.session_state.classified_units[j]
                                    rerun_needed = True
                        with c_excl:
                            if is_excluded:
                                if st.button("복원", key=f"excl_{u['id']}", use_container_width=True):
                                    st.session_state.excluded_unit_ids.remove(u["id"])
                                    rerun_needed = True
                            else:
                                if st.button("✕", key=f"excl_{u['id']}", use_container_width=True):
                                    st.session_state.excluded_unit_ids.append(u["id"])
                                    rerun_needed = True

            st.markdown('</div>', unsafe_allow_html=True)

            # Manual text addition
            st.markdown(
                '<div class="card" style="margin-top:8px;">'
                '<div class="card-title" style="margin-bottom:8px;">텍스트 직접 추가</div>',
                unsafe_allow_html=True,
            )
            manual_text = st.text_input(
                "한국어 텍스트",
                key=f"manual_text_{active_slide}",
                placeholder="인식되지 않은 텍스트를 입력하세요",
                label_visibility="collapsed",
            )
            mc_cat, mc_add = st.columns([2, 1])
            with mc_cat:
                manual_cat = st.selectbox(
                    "분류",
                    ["발표용", "카피"],
                    key=f"manual_cat_{active_slide}",
                    label_visibility="collapsed",
                )
            with mc_add:
                if st.button("추가", key=f"manual_add_{active_slide}", use_container_width=True, type="primary"):
                    if manual_text.strip():
                        import time as _time
                        new_id = f"manual_s{active_slide}_{int(_time.time()*1000)}"
                        st.session_state.classified_units.append({
                            "id": new_id,
                            "slide_idx": active_slide,
                            "shape_id": -1,
                            "p_idx": 0,
                            "ko_text": manual_text.strip(),
                            "font_size": 14.0,
                            "shape_text": manual_text.strip(),
                            "shape_para_count": 1,
                            "category": "presentation" if manual_cat == "발표용" else "copy",
                        })
                        rerun_needed = True
            st.markdown('</div>', unsafe_allow_html=True)

        # Nav buttons outside the scrollable container — always visible
        if rerun_needed:
            st.rerun()

        c_back, c_next = st.columns(2)
        with c_back:
            if st.button("← 이전", key="cl_back", use_container_width=True):
                st.session_state.stage = "upload"
                st.rerun()
        with c_next:
            if st.button("발표용 감수 →", key="cl_next", type="primary", use_container_width=True):
                _excl_ids = set(st.session_state.excluded_unit_ids)
                pres = [u for u in st.session_state.classified_units
                        if u["category"] == "presentation" and u["id"] not in _excl_ids]
                copy = [u for u in st.session_state.classified_units
                        if u["category"] == "copy" and u["id"] not in _excl_ids]
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
            '<div class="review-img-panel">'
            '<div class="bezel">',
            unsafe_allow_html=True,
        )
        st.markdown(_slide_img_html(slide_idx), unsafe_allow_html=True)
        st.markdown(
            f'<div class="bezel-caption">슬라이드 {slide_idx + 1}</div>'
            '</div></div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<div style="font-size:11px;color:#9CA3AF;text-align:center;margin-top:4px;">'
            '미리보기는 서버 폰트 제한으로 실제 PPT와 다를 수 있습니다'
            '</div>',
            unsafe_allow_html=True,
        )

    with col_panel:
        with st.container(height=650):
            # Check if translations are actually populated — show retry if empty
            _n_translated = sum(1 for u in slide_units if trans.get(u["id"], {}).get("en_text", "").strip())
            if len(slide_units) > 0 and _n_translated == 0:
                st.warning(f"번역 결과가 없습니다 (0/{len(slide_units)}). 번역을 다시 실행해주세요.")
                if st.button("번역 재실행", key="retry_trans", type="primary"):
                    st.session_state.translations_loaded = False
                    st.rerun()
            for unit in slide_units:
                item = trans.get(unit["id"], {})
                en_text = item.get("en_text", "")
                notes = item.get("notes", "") or item.get("clarification", "")
                uid = unit["id"]

                with st.container(border=True, key=f"pair_{uid}"):
                    # Korean source (white block, top)
                    st.markdown(
                        f'<div style="padding:13px 15px;font-size:13.5px;color:var(--ct);'
                        f'line-height:1.7;">{_html.escape(unit["ko_text"])}</div>',
                        unsafe_allow_html=True,
                    )
                    # Editable English translation (navy block, bottom)
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

                if notes:
                    st.markdown(
                        f'<div style="font-size:12px;color:var(--cm);line-height:1.5;'
                        f'padding:0 4px;margin:-6px 0 14px;">📝 {_html.escape(notes)}</div>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown('<div style="margin-bottom:14px;"></div>', unsafe_allow_html=True)

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
    c_prev, c_counter, c_next_btn = st.columns([1, 1.5, 1])
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

    _has_copy = bool(st.session_state.copy_units)
    _n_copy = len(st.session_state.copy_units)
    _next_label = f"다음 단계: 카피 선택 ({_n_copy}개) →" if _has_copy else "다음 단계: 다운로드 →"
    c_2a_back, c_2a_fwd = st.columns([1, 2])
    with c_2a_back:
        if st.button("← 분류로 돌아가기", key="2a_back", use_container_width=True):
            st.session_state.stage = "classify"
            st.rerun()
    with c_2a_fwd:
        if st.button(_next_label, key="2a_done", type="primary", use_container_width=True):
            if _has_copy:
                st.session_state.copy_options_loaded = False
                st.session_state.stage = "review_2b"
            else:
                st.session_state.stage = "download"
            st.rerun()


# ── Stage: review_2b ──────────────────────────────────────────────────────────
elif st.session_state.stage == "review_2b":

    copy_units = st.session_state.copy_units

    if not st.session_state.copy_options_loaded:
        n_copy = len(copy_units)
        st.markdown(
            f'<div style="padding:24px 32px 8px;">',
            unsafe_allow_html=True,
        )
        with st.spinner(f"AI가 카피 옵션을 생성하는 중... ({n_copy}개 카피 항목)"):
            try:
                opts = generate_copy_options(copy_units, _build_context())
                st.session_state.copy_options = opts
                st.session_state.copy_selections = {
                    u["id"]: opts.get(u["id"], {}).get("options", [""])[0]
                    for u in copy_units
                }
                st.session_state.copy_options_loaded = True
            except Exception as e:
                st.error(f"카피 옵션 생성 중 오류가 발생했습니다: {e}")
                st.markdown(
                    '<div style="margin-top:8px;font-size:13px;color:#667085;">' +
                    f'카피 항목 수: {n_copy}개' +
                    '</div>',
                    unsafe_allow_html=True,
                )
                if st.button("다시 시도", key="retry_copy", type="primary"):
                    st.rerun()
                st.stop()
        st.rerun()

    # Group duplicate ko_text — user selects once for all occurrences
    _ko_groups: dict = {}
    for u in copy_units:
        _ko_groups.setdefault(u["ko_text"], []).append(u)
    unique_groups = list(_ko_groups.values())

    total = len(unique_groups)
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

    group = unique_groups[idx]
    unit = group[0]  # representative unit for copy_options lookup
    unit_data = st.session_state.copy_options.get(unit["id"], {})
    options = unit_data.get("options", ["", "", ""])
    notes = unit_data.get("notes", "")
    recommendation = unit_data.get("recommendation", "")
    cultural_flag = unit_data.get("cultural_flag", "")
    clarification = unit_data.get("clarification", "")
    current_sel = st.session_state.copy_selections.get(unit["id"], options[0] if options else "")
    slide_idx = unit["slide_idx"]
    dup_slides = sorted({u["slide_idx"] for u in group})

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
        _cap_extra = ""
        if len(dup_slides) > 1:
            _slides_str = ", ".join(str(s + 1) for s in dup_slides)
            _cap_extra = f' (슬라이드 {_slides_str}에 반복)'
        st.markdown(
            f'<div class="bezel-caption">슬라이드 {slide_idx + 1}{_cap_extra}</div>'
            '</div></div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<div style="font-size:11px;color:#9CA3AF;text-align:center;margin-top:4px;">'
            '미리보기는 서버 폰트 제한으로 실제 PPT와 다를 수 있습니다'
            '</div>',
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

        # Duplicate notice
        if len(dup_slides) > 1:
            _slides_str = ", ".join(str(s + 1) for s in dup_slides)
            st.markdown(
                f'<div style="font-size:12px;color:#667085;border-radius:6px;'
                f'padding:6px 12px;margin-bottom:10px;background:#F9F5FF;'
                f'border:1px solid #D6BBFB;">'
                f'🔁 슬라이드 {_slides_str}에 동일 카피 — 한 번 선택하면 모두 적용됩니다</div>',
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
                    for _gu in group:
                        st.session_state.copy_selections[_gu["id"]] = opt_text
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

        # Manual edit section
        st.markdown(
            '<div style="margin-top:12px;border-top:1px solid #E5E7EB;padding-top:12px;">'
            '<div style="font-size:12.5px;color:#6B7280;font-weight:500;margin-bottom:6px;">직접 수정</div>',
            unsafe_allow_html=True,
        )
        manual_val = st.text_area(
            "직접 수정",
            value=current_sel,
            key=f"manual_copy_{unit['id']}",
            height=70,
            label_visibility="collapsed",
        )
        mc_use, mc_check = st.columns(2)
        with mc_use:
            if st.button("이 카피 사용", key=f"use_manual_{unit['id']}", use_container_width=True, type="primary"):
                for _gu in group:
                    st.session_state.copy_selections[_gu["id"]] = manual_val
                st.rerun()
        with mc_check:
            if st.button("문법 체크", key=f"grammar_{unit['id']}", use_container_width=True):
                with st.spinner("문법 체크 중..."):
                    try:
                        feedback = check_copy_grammar(unit["ko_text"], manual_val, _build_context())
                        st.session_state.copy_grammar_results[unit["id"]] = feedback
                    except Exception as e:
                        st.session_state.copy_grammar_results[unit["id"]] = f"오류: {e}"
                st.rerun()
        grammar_feedback = st.session_state.copy_grammar_results.get(unit["id"], "")
        if grammar_feedback:
            st.markdown(
                f'<div style="background:#F0FDF4;border:1px solid #86EFAC;border-radius:8px;'
                f'padding:10px 14px;font-size:13px;color:#166534;line-height:1.6;margin-top:8px;">'
                f'✅ {_html.escape(grammar_feedback)}</div>',
                unsafe_allow_html=True,
            )
        st.markdown('</div>', unsafe_allow_html=True)

        # AI chat section
        st.markdown(
            '<div style="margin-top:10px;border-top:1px solid #E5E7EB;padding-top:10px;">'
            '<div style="font-size:12.5px;color:#6B7280;font-weight:500;margin-bottom:6px;">AI 수정 요청</div>',
            unsafe_allow_html=True,
        )
        _chat_col, _chat_btn = st.columns([5, 1])
        with _chat_col:
            user_msg = st.text_input(
                "AI 수정 요청",
                key=f"chat_2b_{idx}",
                placeholder=f"수정 요청 내용을 입력하세요 ({idx+1}/{total})",
                label_visibility="collapsed",
            )
        with _chat_btn:
            _send = st.button("전송", key=f"send_2b_{idx}", use_container_width=True, type="primary")
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)

    if _send and user_msg.strip():
        with st.spinner("카피 수정 중..."):
            try:
                refined = chat_refine_copy(unit["ko_text"], current_sel, user_msg.strip(), _build_context())
                for _gu in group:
                    st.session_state.copy_selections[_gu["id"]] = refined
            except Exception as e:
                st.error(f"오류: {e}")
        st.rerun()

    # ── Navigation: prev/next copy ──────────────────────────────────────────
    c_prev, c_counter, c_next_btn = st.columns([1, 1.5, 1])
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

    # ── Navigation: stage back/forward ──────────────────────────────────────
    c_2b_back, c_2b_fwd = st.columns([1, 2])
    with c_2b_back:
        if st.button("← 발표용 감수로", key="2b_back", use_container_width=True):
            _back_stage = "review_2a" if st.session_state.presentation_units else "classify"
            st.session_state.stage = _back_stage
            st.rerun()
    with c_2b_fwd:
        if st.button("다음 단계: 다운로드 →", key="2b_done", type="primary", use_container_width=True):
            st.session_state.stage = "download"
            st.rerun()


# ── Stage: en_ko (영→한 번역) ─────────────────────────────────────────────────
elif st.session_state.stage == "en_ko":

    text_units = st.session_state.text_units

    if not st.session_state.en_ko_loaded:
        n = len(text_units)
        with st.spinner(f"AI가 영어 텍스트를 한국어로 번역하는 중... ({n}개 항목)"):
            try:
                ko_trans = translate_en_to_ko(text_units, _build_context())
                st.session_state.en_ko_translations = ko_trans
                st.session_state.en_ko_loaded = True
            except Exception as e:
                st.error(f"번역 오류: {e}")
                if st.button("다시 시도", key="retry_en_ko", type="primary"):
                    st.rerun()
                st.stop()
        st.rerun()

    ko_trans = st.session_state.en_ko_translations
    n_total = len(text_units)
    n_done = sum(1 for v in ko_trans.values() if v.strip())

    st.markdown('<div class="page">', unsafe_allow_html=True)
    st.markdown('<div class="page-title">번역 검토</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="page-sub">총 {n_total}개 텍스트 중 {n_done}개 번역 완료</div>',
        unsafe_allow_html=True,
    )

    # Preview table: English original → Korean translation
    st.markdown('<div class="card">', unsafe_allow_html=True)
    rows = []
    for u in text_units:
        en = u["ko_text"]  # source stored in ko_text field
        ko = ko_trans.get(u["id"], "")
        if en.strip():
            rows.append({"영어 원문": en, "한국어 번역": ko})
    if rows:
        import pandas as _pd
        st.dataframe(
            _pd.DataFrame(rows),
            use_container_width=True,
            hide_index=True,
            height=min(400, 40 + 35 * len(rows)),
        )
    st.markdown('</div>', unsafe_allow_html=True)

    c_back, c_next = st.columns([1, 2])
    with c_back:
        if st.button("← 업로드로 돌아가기", key="en_ko_back", use_container_width=True):
            st.session_state.stage = "upload"
            st.rerun()
    with c_next:
        if st.button("한국어 파일 생성 및 다운로드 →", key="en_ko_next", type="primary", use_container_width=True):
            st.session_state.output_bytes = None
            st.session_state.stage = "download"
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)


# ── Stage: download ───────────────────────────────────────────────────────────
elif st.session_state.stage == "download":

    _is_en_ko = st.session_state.direction == "en_ko"
    _file_type = st.session_state.file_type

    if st.session_state.output_bytes is None:
        if _is_en_ko:
            _spinner_msg = f"번역을 {'PPTX' if _file_type == 'pptx' else 'Word 문서'}에 적용하는 중..."
            with st.spinner(_spinner_msg):
                try:
                    if _file_type == "docx":
                        out = apply_translations_to_docx(
                            st.session_state.file_bytes,
                            st.session_state.en_ko_translations,
                        )
                    else:
                        out = apply_translations(
                            st.session_state.file_bytes,
                            st.session_state.en_ko_translations,
                            font_name="",
                        )
                    st.session_state.output_bytes = out
                except Exception as e:
                    st.error(f"파일 생성 오류: {e}")
                    st.stop()
        else:
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

    if _is_en_ko:
        n_total = len(st.session_state.text_units)
        n_done = sum(1 for v in st.session_state.en_ko_translations.values() if v.strip())
        st.markdown(
            f'<div class="page-sub">영어 원문 {n_total}개 항목 → 한국어 번역 {n_done}개 완료</div>',
            unsafe_allow_html=True,
        )
        _mime = (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            if _file_type == "docx"
            else "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
        _dl_label = f"한국어 {'Word' if _file_type == 'docx' else 'PPT'} 다운로드"
    else:
        n_pres = len(st.session_state.presentation_units)
        n_copy = len(st.session_state.copy_units)
        st.markdown(
            f'<div class="page-sub">발표용 텍스트 {n_pres}개 번역 · 광고 카피 {n_copy}개 트랜스크리에이션 완료</div>',
            unsafe_allow_html=True,
        )
        _mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        _dl_label = "영문 PPT 다운로드"

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.success("번역본이 준비되었습니다!")
    st.download_button(
        label=_dl_label,
        data=st.session_state.output_bytes,
        file_name=st.session_state.file_name,
        mime=_mime,
        type="primary",
    )
    st.markdown('</div>', unsafe_allow_html=True)

    if st.button("처음으로 돌아가기"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
