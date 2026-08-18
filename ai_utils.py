import json
import os
import re
import anthropic

MODEL = "claude-sonnet-5"


def _client() -> anthropic.Anthropic:
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise ValueError("ANTHROPIC_API_KEY 환경변수가 설정되지 않았습니다.")
    return anthropic.Anthropic(api_key=api_key)


def _extract_json(text: str):
    try:
        return json.loads(text)
    except Exception:
        pass
    for start, end in [("[", "]"), ("{", "}")]:
        s = text.find(start)
        e = text.rfind(end)
        if s >= 0 and e > s:
            try:
                return json.loads(text[s : e + 1])
            except Exception:
                pass
    raise ValueError(f"JSON 파싱 실패: {text[:300]}")


def _glossary_line(glossary: str) -> str:
    return f"\n[Glossary — Must NOT translate or alter]: {glossary}" if glossary.strip() else ""


def classify_text_units(text_units: list[dict], glossary: str) -> list[dict]:
    """Add 'category': 'presentation' | 'copy' to each unit."""
    client = _client()
    input_list = [{"id": u["id"], "ko_text": u["ko_text"]} for u in text_units]

    prompt = f"""You are analyzing Korean presentation slide text.

Classify each text unit as:
- "presentation": titles, bullet points, data labels, informational text
- "copy": advertising slogans, taglines, brand messages, emotional/creative marketing text
{_glossary_line(glossary)}

Input:
{json.dumps(input_list, ensure_ascii=False, indent=2)}

Respond with ONLY a JSON array:
[{{"id": "...", "category": "presentation"}}]"""

    resp = client.messages.create(
        model=MODEL,
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )
    result = _extract_json(resp.content[0].text)
    cat_map = {item["id"]: item["category"] for item in result}
    return [{**u, "category": cat_map.get(u["id"], "presentation")} for u in text_units]


def translate_presentation_texts(
    presentation_units: list[dict], glossary: str
) -> dict[str, str]:
    """Batch-translate all presentation units. Returns id → en_text."""
    client = _client()
    input_list = [{"id": u["id"], "ko_text": u["ko_text"]} for u in presentation_units]

    prompt = f"""You are an expert presentation translator localizing Korean slides for a US audience.
Translate each text to natural, professional American English — concise and faithful.
{_glossary_line(glossary)}

Input:
{json.dumps(input_list, ensure_ascii=False, indent=2)}

Respond with ONLY a JSON array:
[{{"id": "...", "en_text": "..."}}]"""

    resp = client.messages.create(
        model=MODEL,
        max_tokens=8192,
        messages=[{"role": "user", "content": prompt}],
    )
    result = _extract_json(resp.content[0].text)
    return {item["id"]: item["en_text"] for item in result}


def generate_copy_options(
    copy_units: list[dict], glossary: str
) -> dict[str, dict]:
    """Generate 3 transcreation options per copy unit with rich annotations.

    Returns id -> {
        options: [creative, balanced, faithful],
        notes: nuance/rhyme/wordplay explanation,
        recommendation: which option and why,
        cultural_flag: US-audience note (empty if n/a),
        clarification: interpretation assumption if ambiguous (empty if clear)
    }
    """
    client = _client()
    input_list = [{"id": u["id"], "ko_text": u["ko_text"]} for u in copy_units]

    prompt = f"""You are a creative director at an advertising agency, formerly a senior copywriter.

[Context]
Translating Korean TV commercial (TVC) copy for the US market.
- Maintain the content, logical structure, and narrative flow across lines
- Adapt to natural American English; liberal interpretation (의역) is encouraged when it better captures tone, rhythm, or impact
- Pay close attention to rhymes, wordplay, double meanings, and emotional register in the Korean original
{_glossary_line(glossary)}

[Task]
For EACH line, produce:
1. Exactly 3 English options:
   - Option 1 (Creative): Liberal transcreation — prioritize feel, rhythm, punch, and US market resonance
   - Option 2 (Balanced): Balance faithfulness to meaning with natural American English flow
   - Option 3 (Faithful): Closest to the literal Korean meaning, minimal interpretation
2. "notes": In English, explain any rhymes, puns, double meanings, emotional tone, or stylistic features in the Korean original that informed your choices
3. "recommendation": State which option better preserves original intent vs. which has stronger creative impact, and explain why briefly
4. "cultural_flag": If a Korean cultural element or expression will be opaque to a US audience, explain the issue and suggest how to handle it. Empty string if not applicable.
5. "clarification": If the Korean meaning is ambiguous or could be interpreted multiple ways, state the interpretation you assumed. Empty string if meaning is clear.

Consider the full sequence together — maintain campaign/narrative flow across all lines.

[Input]
{json.dumps(input_list, ensure_ascii=False, indent=2)}

[Output Format]
Respond with ONLY a JSON array:
[
  {{
    "id": "...",
    "options": ["Option 1 creative", "Option 2 balanced", "Option 3 faithful"],
    "notes": "...",
    "recommendation": "...",
    "cultural_flag": "...",
    "clarification": "..."
  }}
]"""

    resp = client.messages.create(
        model=MODEL,
        max_tokens=8192,
        messages=[{"role": "user", "content": prompt}],
    )
    result = _extract_json(resp.content[0].text)
    return {
        item["id"]: {
            "options": item.get("options", ["", "", ""]),
            "notes": item.get("notes", ""),
            "recommendation": item.get("recommendation", ""),
            "cultural_flag": item.get("cultural_flag", ""),
            "clarification": item.get("clarification", ""),
        }
        for item in result
    }


def chat_modify_presentation(
    presentation_units: list[dict],
    current_translations: dict[str, str],
    instruction: str,
    glossary: str,
) -> dict[str, str]:
    """Apply user instruction to all presentation translations. Returns updated id → en_text."""
    client = _client()
    state = [
        {
            "id": u["id"],
            "ko_text": u["ko_text"],
            "en_text": current_translations.get(u["id"], ""),
        }
        for u in presentation_units
    ]

    prompt = f"""You are a presentation translation editor.

Current translations:
{json.dumps(state, ensure_ascii=False, indent=2)}
{_glossary_line(glossary)}

User instruction: "{instruction}"

Apply the instruction and return ALL translations (updated + unchanged).

Respond with ONLY a JSON array:
[{{"id": "...", "en_text": "..."}}]"""

    resp = client.messages.create(
        model=MODEL,
        max_tokens=8192,
        messages=[{"role": "user", "content": prompt}],
    )
    result = _extract_json(resp.content[0].text)
    return {item["id"]: item["en_text"] for item in result}


def chat_refine_copy(
    ko_text: str,
    current_en: str,
    instruction: str,
    glossary: str,
) -> str:
    """Refine a single TVC copy line based on user instruction. Returns refined English text."""
    client = _client()

    prompt = f"""You are a creative director at an advertising agency, formerly a senior copywriter.
You are refining English TVC copy translated from Korean for the US market.

Korean original: "{ko_text}"
Current English version: "{current_en}"
{_glossary_line(glossary)}

User instruction: "{instruction}"

Apply the instruction while preserving the nuance, rhythm, and intent of the Korean original.
Reply with ONLY the refined English text — no quotes, no explanation, no preamble."""

    resp = client.messages.create(
        model=MODEL,
        max_tokens=512,
        messages=[{"role": "user", "content": prompt}],
    )
    return resp.content[0].text.strip().strip('"')
