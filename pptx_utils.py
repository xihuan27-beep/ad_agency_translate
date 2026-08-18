import io
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import MSO_AUTO_SIZE


def extract_text_units(file_bytes: bytes) -> list[dict]:
    prs = Presentation(io.BytesIO(file_bytes))
    text_units = []
    for s_idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for p_idx, para in enumerate(shape.text_frame.paragraphs):
                text = para.text.strip()
                if len(text) < 5:
                    continue
                font_size = 14.0
                if para.runs and para.runs[0].font.size:
                    font_size = para.runs[0].font.size.pt
                text_units.append({
                    "id": f"s{s_idx}_sh{shape.shape_id}_p{p_idx}",
                    "slide_idx": s_idx,
                    "shape_id": shape.shape_id,
                    "p_idx": p_idx,
                    "ko_text": text,
                    "font_size": font_size,
                })
    return text_units


def apply_translations(file_bytes: bytes, translations: dict[str, str]) -> bytes:
    prs = Presentation(io.BytesIO(file_bytes))

    for s_idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            tf = shape.text_frame
            orig_auto_size = tf.auto_size
            frame_modified = False

            for p_idx, para in enumerate(tf.paragraphs):
                unit_id = f"s{s_idx}_sh{shape.shape_id}_p{p_idx}"
                if unit_id not in translations:
                    continue
                en_text = translations[unit_id]
                ko_len = len(para.text)
                en_len = len(en_text)
                if ko_len == 0 or en_len == 0:
                    continue

                run_sizes = [r.font.size.pt if r.font.size else None for r in para.runs]

                for run in para.runs:
                    run.text = ""
                if not para.runs:
                    continue
                para.runs[0].text = en_text

                # Korean chars are ~1.3× average English char width
                effective_ko_width = ko_len * 1.3
                if en_len > effective_ko_width:
                    scale = max(effective_ko_width / en_len, 0.65)
                    orig_size = run_sizes[0] if run_sizes else None
                    if orig_size:
                        para.runs[0].font.size = Pt(max(orig_size * scale, 8.0))

                frame_modified = True

            # Enable PowerPoint-native auto-fit, but leave SHAPE_TO_FIT_TEXT frames alone
            if frame_modified and orig_auto_size != MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT:
                try:
                    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                except Exception:
                    pass

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()


def extract_reference_texts(file_bytes: bytes) -> list[str]:
    """Extract all non-empty text lines from a reference (already-translated) PPTX."""
    prs = Presentation(io.BytesIO(file_bytes))
    texts = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    if len(text) >= 3:
                        texts.append(text)
    return texts
