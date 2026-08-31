def build_context(key_phrases: list[dict] | None = None) -> str:
    """Build the AI prompt context string from the term-pair list collected on
    the upload page. Mirrors the Streamlit app's _build_context(), minus the
    brand-name/reference-pptx/glossary fields that were removed from the UI.
    """
    parts = []
    kp = [
        p for p in (key_phrases or [])
        if p.get("한국어", "").strip() and p.get("영어", "").strip()
    ]
    if kp:
        lines = "\n".join(f"  '{p['한국어']}' → '{p['영어']}'" for p in kp)
        parts.append(f"Preferred term translations (use these exact English expressions):\n{lines}")
    return "\n\n".join(parts)
