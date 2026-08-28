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
  --ct3: #8A9BB8;   /* tertiary / faint text */
  --cb:  #DCE3F0;   /* border */
  --cbf: #EDF0F8;   /* faint border (hairlines) */
  --cw:  #FFFFFF;
  --cbl: #E8ECFA;   /* light navy tint */
  --cnt: #F3F5FD;   /* navy tint (very light) */
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
  height: 58px; background: var(--cw); border-bottom: 1px solid var(--cb);
}
.topbar-inner {
  max-width: 1200px; margin: 0 auto; height: 100%; padding: 0 32px;
  display: flex; align-items: center; justify-content: center; gap: 12px;
}
.topbar-logo {
  width: 32px; height: 32px; border-radius: 8px; background: var(--cp);
  display: flex; align-items: center; justify-content: center; flex-shrink: 0;
}
.topbar-title { font-size: 19px; font-weight: 700; color: var(--cp); letter-spacing: -0.02em; }

/* ── Step rail ── */
.steprail {
  border-bottom: 1px solid var(--cb); background: var(--cw); height: 48px;
}
.steprail-inner {
  max-width: 1200px; margin: 0 auto; height: 100%; padding: 0 28px;
  display: flex; align-items: stretch; justify-content: center;
}
.step {
  display: flex; align-items: center; gap: 7px; padding: 0 16px;
  font-size: 13px; font-weight: 500; color: var(--cm);
  position: relative; white-space: nowrap;
}
.step.active { color: var(--cp); font-weight: 600; }
.step.active::after {
  content: ''; position: absolute; bottom: 0; left: 0; right: 0;
  height: 2px; background: var(--cp); border-radius: 2px 2px 0 0;
}
.step-n {
  width: 7px; height: 7px; border-radius: 50%;
  flex-shrink: 0; background: var(--cb); transition: background .15s;
}
.step.active .step-n { background: var(--cp); }
.step.done .step-n   { background: var(--cm); }
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
  background: #18182A; border-radius: 10px; padding: 4px; margin-bottom: 0;
  box-shadow: 0 4px 20px rgba(0,0,0,.2);
}
.bezel img { border-radius: 7px; width: 100%; display: block; }
.bezel-placeholder {
  border-radius: 7px; background: var(--cw); width: 100%; aspect-ratio: 16/9;
  display: flex; align-items: center; justify-content: center;
  color: var(--ct3); font-size: 14px;
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
div.stButton > button[kind="primary"],
div.stDownloadButton > button[kind="primary"] {
  background: var(--cp) !important; border-color: var(--cp) !important; color: #fff !important;
}
div.stButton > button[kind="primary"]:hover,
div.stDownloadButton > button[kind="primary"]:hover {
  background: var(--cpa) !important; border-color: var(--cpa) !important;
}
div.stButton > button:disabled,
div.stButton > button[kind="primary"]:disabled,
div.stDownloadButton > button:disabled {
  background: var(--cbg) !important; border-color: var(--cb) !important;
  color: var(--ct3) !important; opacity: 1 !important; cursor: default !important;
}

/* ── Input fields ── */
.stTextInput > div > div > input,
.stTextArea > div > div > textarea {
  background: #fff !important; color: var(--ct) !important;
  border: 1px solid var(--cb) !important; border-radius: 7px !important;
  font-size: 14px !important; box-shadow: none !important;
}
.stTextInput > div > div > input::placeholder,
.stTextArea > div > div > textarea::placeholder { color: var(--ct3) !important; }
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
/* Item-card tag buttons sit in a narrow column — keep the label on one line */
[class*="st-key-tog_pres_"] button, [class*="st-key-tog_copy_"] button {
  white-space: nowrap !important; font-size: 11.5px !important; padding: 4px 6px !important;
}

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

/* ── Copy option card: tag above an inline-editable box, no star ── */
[class*="st-key-copyopt_"] {
  background: var(--cw) !important; border-radius: 8px !important; box-shadow: var(--cs) !important;
  padding: 13px 15px !important; margin-bottom: 8px !important; border: 1px solid var(--cb) !important;
}
[class*="st-key-copyopt_"] textarea {
  border: none !important; box-shadow: none !important; background: transparent !important;
  color: var(--ct) !important; font-size: 14.5px !important; font-weight: 500 !important;
  padding: 0 !important; font-family: var(--font) !important;
}
[class*="st-key-copyopt_sel_"] {
  background: var(--cp) !important; border-color: var(--cp) !important;
}
[class*="st-key-copyopt_sel_"] textarea,
[class*="st-key-copyopt_sel_"] [data-baseweb="textarea"],
[class*="st-key-copyopt_sel_"] [data-testid="stTextAreaRootElement"] { color: #fff !important; background: transparent !important; }
[class*="st-key-copyopt_sel_"] .cr-lname {
  background: rgba(255,255,255,.14) !important; color: rgba(255,255,255,.85) !important; border-color: rgba(255,255,255,.25) !important;
}
[class*="st-key-copyopt_sel_"] .cr-lsub { color: rgba(255,255,255,.65) !important; }

/* ── 발표용 감수 / 카피 선택: centered content with side margins, not edge-to-edge ── */
[class*="st-key-review_2a_wrap"],
[class*="st-key-review_2b_wrap"] {
  max-width: 1200px !important; margin: 0 auto !important; padding: 0 40px !important;
}

/* ── Classify page: top nav row (back | description | forward), centered ── */
[class*="st-key-classify_navrow"] {
  max-width: 1200px !important; margin: 0 auto !important;
  padding: 11px 40px !important; border-bottom: 1px solid var(--cbf) !important;
  margin-bottom: 14px !important;
}

/* ── Classify item merge states: shared pale block, seam edges squared off ── */
[class*="st-key-item_mtop_"] {
  background: var(--cbl) !important; border: none !important; box-shadow: none !important;
  border-radius: 8px 8px 0 0 !important; padding: 11px 13px !important; margin-bottom: 0 !important;
}
[class*="st-key-item_mmid_"] {
  background: var(--cbl) !important; border: none !important; box-shadow: none !important;
  border-radius: 0 !important; padding: 11px 13px !important; margin-bottom: 0 !important;
}
[class*="st-key-item_mbot_"] {
  background: var(--cbl) !important; border: none !important; box-shadow: none !important;
  border-radius: 0 0 8px 8px !important; padding: 11px 13px !important; margin-bottom: 8px !important;
}

/* ── Separator between classify items: dashed line + circular toggle ── */
.sep-line-half { height: 0; border-top: 1px dashed var(--cb); margin-top: 11px; }
.sep-line-half.faint { border-top-color: var(--cbl); opacity: .7; }
[class*="st-key-sepbtn_"] button {
  border-radius: 50% !important; width: 22px !important; height: 22px !important;
  min-height: 22px !important; padding: 0 !important; font-size: 10px !important;
  font-weight: 700 !important; line-height: 1 !important;
}
[class*="st-key-sepbtn_merged_"] button {
  background: var(--cbl) !important; color: var(--cp) !important;
  border-style: dashed !important; border-color: rgba(12,39,144,.3) !important;
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
        "fetched_url": "",
        "fetched_file_bytes": None,
        "fetched_file_type": "pptx",
        "merged_seps": {},  # slide_idx -> set of positions merged with the next item
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


def _materialize_merges(units: list[dict], merged_seps: dict) -> list[dict]:
    """Combine consecutive items joined via the classify page's merge toggle into
    one translatable unit each, without mutating the reversible classify state."""
    by_slide = defaultdict(list)
    for u in units:
        by_slide[u["slide_idx"]].append(u)
    result = []
    for slide_idx, items in by_slide.items():
        seps = merged_seps.get(slide_idx, set())
        n = len(items)
        pos = 0
        while pos < n:
            run = [items[pos]]
            while pos in seps and pos + 1 < n:
                pos += 1
                run.append(items[pos])
            if len(run) == 1:
                result.append(run[0])
            else:
                merged_unit = dict(run[0])
                merged_unit["ko_text"] = " ".join(r["ko_text"] for r in run)
                merged_unit["shape_text"] = merged_unit["ko_text"]
                result.append(merged_unit)
            pos += 1
    return result


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
LABELS_KO_EN = ["파일 업로드", "슬라이드 분류", "발표용 감수", "카피 선택", "다운로드"]
STAGES_EN_KO = ["upload", "en_ko", "download"]
LABELS_EN_KO = ["파일 업로드", "번역", "다운로드"]

def _current_stages() -> tuple[list, list]:
    if st.session_state.direction == "en_ko":
        return STAGES_EN_KO, LABELS_EN_KO
    return STAGES_KO_EN, LABELS_KO_EN

def _render_chrome():
    stages, labels = _current_stages()
    curr_stage = st.session_state.stage
    curr_idx = stages.index(curr_stage) if curr_stage in stages else 0
    st.markdown(
        '<div class="topbar"><div class="topbar-inner">'
        '<div class="topbar-logo">'
        '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">'
        '<rect x="2" y="3" width="20" height="15" rx="2" stroke="white" stroke-width="2" fill="none"/>'
        '<path d="M8 22h8M12 18v4" stroke="white" stroke-width="2" stroke-linecap="round"/>'
        '</svg></div>'
        '<span class="topbar-title">Agency Deck Translator</span>'
        '</div></div>',
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
        items += f'<div class="{cls}"><span class="step-n"></span>{label}</div>'
    st.markdown(f'<div class="steprail"><div class="steprail-inner">{items}</div></div>', unsafe_allow_html=True)

_render_chrome()


# ── Stage: upload ─────────────────────────────────────────────────────────────
if st.session_state.stage == "upload":
    st.markdown('<div class="page">', unsafe_allow_html=True)

    # ── 소제목 1: 번역 선택 ──────────────────────────────────────────────────
    st.markdown('<div class="section-eyebrow">번역 선택</div>', unsafe_allow_html=True)
    _dir_col1, _dir_col2 = st.columns(2)
    with _dir_col1:
        if st.button("한국어 → 영어", key="dir_ko_en", use_container_width=True,
                     type="primary" if st.session_state.direction == "ko_en" else "secondary"):
            if st.session_state.direction != "ko_en":
                st.session_state.direction = "ko_en"
                st.rerun()
    with _dir_col2:
        if st.button("영어 → 한국어", key="dir_en_ko", use_container_width=True,
                     type="primary" if st.session_state.direction == "en_ko" else "secondary"):
            if st.session_state.direction != "en_ko":
                st.session_state.direction = "en_ko"
                st.rerun()

    _is_en_ko = st.session_state.direction == "en_ko"

    # ── 소제목 2: 번역할 PPT 파일 업로드 하기 ──────────────────────────────────
    with st.container(border=True, key="card_file"):
        st.markdown('<div class="card-title">번역할 PPT 파일 업로드 하기</div>', unsafe_allow_html=True)
        c1, c2 = st.columns([5, 1])
        with c1:
            drive_url = st.text_input(
                "drive_url", placeholder="https://drive.google.com/file/d/...",
                label_visibility="collapsed", key="up_drive_url",
            )
        _url_stripped = drive_url.strip()
        _already_fetched = bool(_url_stripped) and _url_stripped == st.session_state.fetched_url
        with c2:
            fetch_clicked = st.button(
                "가져오기", key="btn_fetch_main", use_container_width=True,
                type="secondary" if _already_fetched else "primary",
                disabled=_already_fetched or not _url_stripped,
            )
        st.markdown(
            '<div style="font-size:11.5px;color:var(--cm);margin-top:8px;line-height:1.7;">'
            '무료버전에서는 Google Drive만 가능합니다<br>'
            '링크 공유 시 권한을 편집자(edit)로 해야 합니다</div>',
            unsafe_allow_html=True,
        )
        if _url_stripped and not _gdrive_file_id(_url_stripped):
            st.warning("올바른 Google Drive / Google Docs 링크가 아닙니다.")

        if fetch_clicked and _url_stripped and _gdrive_file_id(_url_stripped):
            fid = _gdrive_file_id(_url_stripped)
            _is_slides = _gdrive_is_slides(_url_stripped)
            _is_docs = _gdrive_is_docs(_url_stripped)
            with st.spinner("Google Drive에서 파일 다운로드 중..."):
                try:
                    _fetched_bytes = _download_gdrive(fid, is_slides=_is_slides, is_docs=_is_docs)
                except Exception as e:
                    st.error(f"다운로드 실패: {e}")
                    st.stop()
            if not _is_zip(_fetched_bytes):
                st.error(
                    "파일을 올바르게 다운로드하지 못했습니다. "
                    "Google Drive 파일 공유 설정을 확인해 주세요:\n\n"
                    "1. Google Drive에서 파일 우클릭 → 공유\n"
                    "2. '링크가 있는 모든 사용자' 또는 '편집자'로 설정\n"
                    "3. 링크 복사 후 다시 시도"
                )
                st.stop()
            st.session_state.fetched_url = _url_stripped
            st.session_state.fetched_file_bytes = _fetched_bytes
            st.session_state.fetched_file_type = (
                "docx" if (_is_docs or _url_stripped.lower().endswith(".docx")) else "pptx"
            )
            st.rerun()

    # ── 소제목 3: 번역 퀄리티 상승을 위한 추가 정보 입력 ───────────────────────
    with st.container(border=True, key="card_quality"):
        st.markdown('<div class="card-title">번역 퀄리티 상승을 위한 추가 정보 입력</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="card-sub">브랜드명, 제품명, 내부 약어, 선호 문구 등 주요 용어를 등록하시면 '
            '이를 활용하여 번역하여 검수 업무가 줄어듭니다</div>',
            unsafe_allow_html=True,
        )
        if not st.session_state.key_phrases:
            st.session_state.key_phrases = [{"한국어": "", "영어": ""}]
        _col_order = [("영어", "English"), ("한국어", "한국어")] if _is_en_ko else [("한국어", "한국어"), ("영어", "English")]
        for idx in range(len(st.session_state.key_phrases)):
            c_a, c_b = st.columns(2)
            for col, (lang_key, placeholder) in zip((c_a, c_b), _col_order):
                with col:
                    _val = st.text_input(
                        f"term_{lang_key}_{idx}",
                        value=st.session_state.key_phrases[idx].get(lang_key, ""),
                        placeholder=placeholder, key=f"term_{lang_key}_{idx}",
                        label_visibility="collapsed",
                    )
                    st.session_state.key_phrases[idx][lang_key] = _val
        if st.button("+ 항목 추가", key="add_term_pair"):
            st.session_state.key_phrases.append({"한국어": "", "영어": ""})
            st.rerun()

    # ── 시작하기 ────────────────────────────────────────────────────────────
    file_ready = (
        st.session_state.fetched_file_bytes is not None
        and st.session_state.fetched_url == drive_url.strip()
    )
    _, col_next = st.columns([1, 1])
    with col_next:
        go = st.button("시작하기", type="primary", disabled=not file_ready,
                       use_container_width=True, key="btn_go")

    if go and file_ready:
        file_bytes = st.session_state.fetched_file_bytes
        _file_type = st.session_state.fetched_file_type
        _url = st.session_state.fetched_url
        fid = _gdrive_file_id(_url)

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

        st.session_state.file_type = _file_type
        st.session_state.file_bytes = file_bytes
        _suffix = "KO" if _is_en_ko else "EN"
        st.session_state.file_name = f"gdrive_{fid[:8]}_{_suffix}.{_file_type}"
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

    # Top nav row: back | description | forward
    with st.container(key="classify_navrow"):
        nc1, nc2, nc3 = st.columns([1, 3, 1])
        with nc1:
            if st.button("← 업로드로 돌아가기", key="cl_back", use_container_width=True):
                st.session_state.stage = "upload"
                st.rerun()
        with nc2:
            st.markdown(
                '<div style="text-align:center;font-size:12.5px;color:var(--cm);padding-top:9px;">'
                'Copywriter가 번역해야할 카피와 아닌 발표용을 구분할 수 있습니다</div>',
                unsafe_allow_html=True,
            )
        with nc3:
            if st.button("발표용 감수로 넘어가기 →", key="cl_next", type="primary", use_container_width=True):
                merged_units = _materialize_merges(st.session_state.classified_units, st.session_state.merged_seps)
                _excl_ids = set(st.session_state.excluded_unit_ids)
                pres = [u for u in merged_units if u["category"] == "presentation" and u["id"] not in _excl_ids]
                copy = [u for u in merged_units if u["category"] == "copy" and u["id"] not in _excl_ids]
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
            '<div style="font-size:11px;color:var(--ct3);text-align:center;margin-top:6px;">'
            '미리보기는 서버 폰트 제한으로 실제 PPT와 다를 수 있습니다'
            '</div>',
            unsafe_allow_html=True,
        )

    # Right: current slide text units (independently scrollable); nav buttons below
    with col_right:
        with st.container(height=650):
            slide_items = slide_groups.get(active_slide, [])
            with st.container(border=True, key=f"card_slide_items_{active_slide}"):
                st.markdown(
                    f'<div class="card-title" style="margin-bottom:10px;">슬라이드 {active_slide + 1}</div>',
                    unsafe_allow_html=True,
                )
                if not slide_items:
                    st.markdown('<div style="font-size:13px;color:var(--cm);">이 슬라이드에 번역할 텍스트가 없습니다.</div>',
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
                    merged_set = st.session_state.merged_seps.get(active_slide, set())
                    for pos, (i, u) in enumerate(slide_items):
                        cat = u["category"]
                        is_excluded = u["id"] in excluded_ids
                        short = u["ko_text"][:45] + ("…" if len(u["ko_text"]) > 45 else "")
                        has_next = pos + 1 < len(slide_items)
                        merge_above = (pos - 1) in merged_set
                        merge_below = pos in merged_set
                        if merge_above and merge_below:
                            _card_key = f"item_mmid_{u['id']}"
                        elif merge_above:
                            _card_key = f"item_mbot_{u['id']}"
                        elif merge_below:
                            _card_key = f"item_mtop_{u['id']}"
                        else:
                            _card_key = f"item_card_{u['id']}"

                        with st.container(border=True, key=_card_key):
                            c_tag, c_text, c_excl = st.columns([1, 3.5, 0.5])
                            with c_tag:
                                if is_excluded:
                                    st.markdown(
                                        '<p style="font-size:11px;color:var(--ct3);margin:0;padding:5px 0;">제외됨</p>',
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
                                _tc = "var(--ct3)" if is_excluded else "var(--ct)"
                                _td = "line-through" if is_excluded else "none"
                                st.markdown(
                                    f"<p style='font-size:12.5px;color:{_tc};margin:0;"
                                    f"padding:5px 0;line-height:1.4;text-decoration:{_td};'>"
                                    f"{_html.escape(short)}</p>",
                                    unsafe_allow_html=True,
                                )
                            with c_excl:
                                if is_excluded:
                                    if st.button("복원", key=f"excl_{u['id']}", use_container_width=True):
                                        st.session_state.excluded_unit_ids.remove(u["id"])
                                        rerun_needed = True
                                else:
                                    if st.button("✕", key=f"excl_{u['id']}", use_container_width=True):
                                        st.session_state.excluded_unit_ids.append(u["id"])
                                        rerun_needed = True

                        if has_next:
                            is_sep_merged = pos in merged_set
                            _line_cls = "faint" if is_sep_merged else ""
                            sc1, sc2, sc3 = st.columns([1, 0.16, 1])
                            with sc1:
                                st.markdown(f'<div class="sep-line-half {_line_cls}"></div>', unsafe_allow_html=True)
                            with sc2:
                                _sep_key = f"sepbtn_merged_{active_slide}_{pos}" if is_sep_merged else f"sepbtn_{active_slide}_{pos}"
                                if st.button("○" if is_sep_merged else "✕", key=_sep_key,
                                             help="합쳐진 항목 나누기" if is_sep_merged else "다음 항목과 합치기"):
                                    ms = st.session_state.merged_seps.setdefault(active_slide, set())
                                    if is_sep_merged:
                                        ms.discard(pos)
                                    else:
                                        ms.add(pos)
                                    rerun_needed = True
                            with sc3:
                                st.markdown(f'<div class="sep-line-half {_line_cls}"></div>', unsafe_allow_html=True)

            # Manual text addition
            with st.container(border=True, key=f"card_manual_add_{active_slide}"):
                st.markdown('<div class="card-title" style="margin-bottom:8px;">텍스트 직접 추가</div>', unsafe_allow_html=True)
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

        # rerun outside the scrollable container so state updates take effect immediately
        if rerun_needed:
            st.rerun()


# ── Stage: review_2a ──────────────────────────────────────────────────────────
elif st.session_state.stage == "review_2a":
    with st.container(key="review_2a_wrap"):

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
                '<div style="font-size:11px;color:var(--ct3);text-align:center;margin-top:4px;">'
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
                f'color:var(--cm);font-weight:500;">{slide_pos+1}/{total_slides} 슬라이드</div>',
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
    with st.container(key="review_2b_wrap"):

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
                        '<div style="margin-top:8px;font-size:13px;color:var(--cm);">' +
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
                '<div style="font-size:11px;color:var(--ct3);text-align:center;margin-top:4px;">'
                '미리보기는 서버 폰트 제한으로 실제 PPT와 다를 수 있습니다'
                '</div>',
                unsafe_allow_html=True,
            )

        with col_panel:
            st.markdown('<div style="padding:16px 24px 0 8px;">', unsafe_allow_html=True)

            # Korean source text header
            st.markdown(
                f'<div style="background:var(--cw);border:1px solid var(--cb);border-radius:8px;padding:10px 14px;'
                f'font-size:14px;color:var(--ct);line-height:1.55;margin-bottom:12px;">'
                f'{_html.escape(unit["ko_text"])}</div>',
                unsafe_allow_html=True,
            )

            # Duplicate notice
            if len(dup_slides) > 1:
                _slides_str = ", ".join(str(s + 1) for s in dup_slides)
                st.markdown(
                    f'<div style="font-size:12px;color:var(--cp);border-radius:6px;'
                    f'padding:6px 12px;margin-bottom:10px;background:var(--cnt);'
                    f'border:1px solid var(--cbl);">'
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

            # Copy option rows — tag above box, directly editable in place, no star
            rerun_sel = False
            for i, (opt_name, opt_sub) in enumerate(OPTS_META):
                opt_text = options[i] if i < len(options) else ""
                is_sel = (current_sel == opt_text)

                _card_key = f"copyopt_sel_{unit['id']}_{i}" if is_sel else f"copyopt_{unit['id']}_{i}"
                with st.container(border=True, key=_card_key):
                    st.markdown(
                        f'<div class="cr-label"><div class="cr-lname">{opt_name}</div>'
                        f'<div class="cr-lsub">{opt_sub}</div></div>',
                        unsafe_allow_html=True,
                    )
                    new_opt_text = st.text_area(
                        opt_name, value=opt_text, key=f"opt_text_{unit['id']}_{i}",
                        height=70, label_visibility="collapsed",
                    )
                    if new_opt_text != opt_text:
                        was_selected = is_sel
                        options[i] = new_opt_text
                        st.session_state.copy_options[unit["id"]]["options"] = options
                        if was_selected:
                            for _gu in group:
                                st.session_state.copy_selections[_gu["id"]] = new_opt_text
                        st.rerun()
                    if not is_sel:
                        if st.button("이 버전 사용", key=f"use_opt_{unit['id']}_{i}", use_container_width=True):
                            for _gu in group:
                                st.session_state.copy_selections[_gu["id"]] = opt_text
                            rerun_sel = True
                    else:
                        st.markdown(
                            '<div style="text-align:center;font-size:12px;font-weight:600;'
                            'color:#fff;padding:7px 0 0;">✓ 선택됨</div>',
                            unsafe_allow_html=True,
                        )

            if rerun_sel:
                st.rerun()

            # Recommendation box
            if recommendation:
                st.markdown(
                    f'<div class="recbox"><div style="font-size:18px;flex-shrink:0;">📌</div>'
                    f'<div><strong>추천 이유:</strong> {_html.escape(recommendation)}</div></div>',
                    unsafe_allow_html=True,
                )

            if st.button("선택된 카피 문법 체크", key=f"grammar_{unit['id']}"):
                with st.spinner("문법 체크 중..."):
                    try:
                        feedback = check_copy_grammar(unit["ko_text"], current_sel, _build_context())
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

            # AI chat section
            st.markdown(
                '<div style="margin-top:10px;border-top:1px solid var(--cb);padding-top:10px;">'
                '<div style="font-size:12.5px;color:var(--cm);font-weight:500;margin-bottom:6px;">AI 수정 요청</div>',
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
                f'color:var(--cm);font-weight:500;">{idx+1}/{total} 카피</div>',
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
    with st.container(border=True, key="card_en_ko_preview"):
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

    with st.container(border=True, key="card_download"):
        st.success("번역본이 준비되었습니다!")
        st.download_button(
            label=_dl_label,
            data=st.session_state.output_bytes,
            file_name=st.session_state.file_name,
            mime=_mime,
            type="primary",
        )

    if st.button("처음으로 돌아가기"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
