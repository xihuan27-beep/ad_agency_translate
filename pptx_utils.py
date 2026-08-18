import io
from pptx import Presentation
from pptx.util import Pt


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
            for p_idx, para in enumerate(shape.text_frame.paragraphs):
                unit_id = f"s{s_idx}_sh{shape.shape_id}_p{p_idx}"
                if unit_id not in translations:
                    continue

                en_text = translations[unit_id]
                ko_len = len(para.text)
                en_len = len(en_text)

                orig_size = 14.0
                if para.runs and para.runs[0].font.size:
                    orig_size = para.runs[0].font.size.pt

                # Clear all runs then write to first
                for run in para.runs:
                    run.text = ""
                if para.runs:
                    para.runs[0].text = en_text
                    if en_len > ko_len * 1.2:
                        para.runs[0].font.size = Pt(max(orig_size * 0.85, 9.0))

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()
