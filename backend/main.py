import io
import os
import sys

from fastapi import FastAPI, HTTPException, Response
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

_here = os.path.dirname(__file__)
sys.path.insert(0, os.path.join(_here, ".."))
sys.path.insert(0, _here)

from ai_utils import (  # noqa: E402
    classify_text_units,
    translate_presentation_texts,
    generate_copy_options,
    chat_refine_copy,
    check_copy_grammar,
    translate_en_to_ko,
)
from pptx_utils import (  # noqa: E402
    extract_text_units,
    apply_translations,
    render_slides_to_images,
)
from docx_utils import extract_text_from_docx, apply_translations_to_docx  # noqa: E402
from pdf_utils import (  # noqa: E402
    extract_text_from_pdf,
    render_pdf_to_images,
    build_translated_docx,
    page_count as pdf_page_count,
)

from gdrive_utils import (  # noqa: E402
    gdrive_file_id,
    gdrive_is_slides,
    gdrive_is_docs,
    download_gdrive,
)
from context_utils import build_context  # noqa: E402
from session_store import store  # noqa: E402

app = FastAPI(title="Agency Deck Translator API")

_frontend_origin = os.environ.get("FRONTEND_ORIGIN", "*")
app.add_middleware(
    CORSMiddleware,
    allow_origins=[_frontend_origin] if _frontend_origin != "*" else ["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# ── Schemas ────────────────────────────────────────────────────────────────
class FetchRequest(BaseModel):
    url: str


class ClassifyRequest(BaseModel):
    keyPhrases: list[dict] = []


class TranslateUnitsRequest(BaseModel):
    units: list[dict]
    keyPhrases: list[dict] = []


class RefineCopyRequest(BaseModel):
    koText: str
    currentEn: str
    instruction: str
    keyPhrases: list[dict] = []


class GrammarCheckRequest(BaseModel):
    koText: str
    enText: str
    keyPhrases: list[dict] = []


class ApplyRequest(BaseModel):
    translations: dict[str, str]
    fontName: str = ""


# ── Health ───────────────────────────────────────────────────────────────
@app.get("/health")
def health():
    return {"status": "ok"}


# ── Session lifecycle ──────────────────────────────────────────────────────
@app.post("/api/sessions")
def create_session():
    session = store.create()
    return {"sessionId": session.id}


def _get_session_or_404(session_id: str):
    session = store.get(session_id)
    if session is None:
        raise HTTPException(status_code=404, detail="세션을 찾을 수 없습니다. 새로고침 후 다시 시도해주세요.")
    return session


# ── Upload / fetch ──────────────────────────────────────────────────────────
@app.post("/api/sessions/{session_id}/fetch")
def fetch_file(session_id: str, body: FetchRequest):
    session = _get_session_or_404(session_id)

    url = body.url.strip()
    fid = gdrive_file_id(url)
    if not fid:
        raise HTTPException(status_code=400, detail="올바른 Google Drive / Google Docs 링크가 아닙니다.")

    is_slides = gdrive_is_slides(url)
    is_docs = gdrive_is_docs(url)
    try:
        file_bytes = download_gdrive(fid, is_slides=is_slides, is_docs=is_docs)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"다운로드 실패: {e}")

    if file_bytes[:4] == b"PK\x03\x04":
        file_type = "docx" if (is_docs or url.lower().endswith(".docx")) else "pptx"
    elif file_bytes[:5] == b"%PDF-":
        file_type = "pdf"
    else:
        raise HTTPException(
            status_code=422,
            detail=(
                "파일을 올바르게 다운로드하지 못했습니다. Google Drive 파일 공유 설정을 확인해 주세요:\n"
                "1. Google Drive에서 파일 우클릭 → 공유\n"
                "2. '링크가 있는 모든 사용자' 또는 '편집자'로 설정\n"
                "3. 링크 복사 후 다시 시도"
            ),
        )

    if file_type == "docx":
        text_units = extract_text_from_docx(file_bytes)
    elif file_type == "pdf":
        text_units = extract_text_from_pdf(file_bytes)
    else:
        text_units = extract_text_units(file_bytes)

    if not text_units:
        raise HTTPException(status_code=422, detail="번역 가능한 텍스트를 찾지 못했습니다. 파일을 확인해 주세요.")

    slide_images: list[bytes] = []
    slide_count = 0
    if file_type == "pptx":
        slide_images = render_slides_to_images(file_bytes)
        from pptx import Presentation as _Prs
        slide_count = len(_Prs(io.BytesIO(file_bytes)).slides)
    elif file_type == "pdf":
        slide_images = render_pdf_to_images(file_bytes)
        slide_count = pdf_page_count(file_bytes)

    session.file_bytes = file_bytes
    session.file_type = file_type
    session.file_name = f"gdrive_{fid[:8]}.{file_type}"
    session.text_units = text_units
    session.slide_images = slide_images
    session.slide_count = slide_count

    return {
        "fileType": file_type,
        "slideCount": slide_count,
        "hasSlideImages": len(slide_images) > 0,
        "textUnits": text_units,
    }


@app.get("/api/sessions/{session_id}/slide-image/{idx}.png")
def get_slide_image(session_id: str, idx: int):
    session = _get_session_or_404(session_id)
    if idx < 0 or idx >= len(session.slide_images):
        raise HTTPException(status_code=404, detail="슬라이드 이미지를 찾을 수 없습니다.")
    return Response(content=session.slide_images[idx], media_type="image/png")


# ── Classify (AI call, session-scoped text units) ───────────────────────────
@app.post("/api/sessions/{session_id}/classify")
def classify(session_id: str, body: ClassifyRequest):
    session = _get_session_or_404(session_id)
    if not session.text_units:
        raise HTTPException(status_code=400, detail="먼저 파일을 가져와야 합니다.")
    context = build_context(body.keyPhrases)
    try:
        classified = classify_text_units(session.text_units, context)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"분류 오류: {e}")
    return {"units": classified}


# ── Stateless AI translation endpoints ──────────────────────────────────────
@app.post("/api/translate/presentation")
def translate_presentation(body: TranslateUnitsRequest):
    context = build_context(body.keyPhrases)
    try:
        result = translate_presentation_texts(body.units, context)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"번역 오류: {e}")
    return {"translations": result}


@app.post("/api/translate/copy-options")
def translate_copy_options(body: TranslateUnitsRequest):
    context = build_context(body.keyPhrases)
    try:
        result = generate_copy_options(body.units, context)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"카피 옵션 생성 오류: {e}")
    return {"options": result}


@app.post("/api/translate/refine-copy")
def refine_copy(body: RefineCopyRequest):
    context = build_context(body.keyPhrases)
    try:
        refined = chat_refine_copy(body.koText, body.currentEn, body.instruction, context)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"수정 오류: {e}")
    return {"text": refined}


@app.post("/api/translate/grammar-check")
def grammar_check(body: GrammarCheckRequest):
    context = build_context(body.keyPhrases)
    try:
        feedback = check_copy_grammar(body.koText, body.enText, context)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"문법 체크 오류: {e}")
    return {"feedback": feedback}


@app.post("/api/translate/en-to-ko")
def translate_en_to_ko_route(body: TranslateUnitsRequest):
    context = build_context(body.keyPhrases)
    try:
        result = translate_en_to_ko(body.units, context)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"번역 오류: {e}")
    return {"translations": result}


# ── Apply translations & download ───────────────────────────────────────────
@app.post("/api/sessions/{session_id}/apply")
def apply(session_id: str, body: ApplyRequest):
    session = _get_session_or_404(session_id)
    if session.file_bytes is None:
        raise HTTPException(status_code=400, detail="먼저 파일을 가져와야 합니다.")
    try:
        if session.file_type == "docx":
            out = apply_translations_to_docx(session.file_bytes, body.translations)
            session.output_file_name = "translated.docx"
        elif session.file_type == "pdf":
            out = build_translated_docx(session.text_units, body.translations)
            session.output_file_name = "translated.docx"
        else:
            out = apply_translations(session.file_bytes, body.translations, font_name=body.fontName)
            session.output_file_name = "translated.pptx"
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"파일 생성 오류: {e}")
    session.output_bytes = out
    return {"ready": True, "fileName": session.output_file_name}


@app.get("/api/sessions/{session_id}/download")
def download(session_id: str):
    session = _get_session_or_404(session_id)
    if session.output_bytes is None:
        raise HTTPException(status_code=400, detail="아직 생성된 파일이 없습니다.")
    mime = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        if session.file_type in ("docx", "pdf")
        else "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
    return StreamingResponse(
        io.BytesIO(session.output_bytes),
        media_type=mime,
        headers={"Content-Disposition": f'attachment; filename="{session.output_file_name}"'},
    )
