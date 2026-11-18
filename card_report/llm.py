"""선택 기능: 규칙으로 분류 못 한 가맹점만 LLM에 묻는다.

ANTHROPIC_API_KEY가 있으면 Claude, 없고 OPENAI_API_KEY가 있으면 OpenAI를 쓴다.
금액·날짜는 보내지 않고 가맹점명만 보낸다.
"""
from __future__ import annotations

import json
import os

import requests

from .categories import ALL_CATEGORIES

PROMPT = """다음은 한국 신용카드 명세서의 가맹점명 목록입니다.
각 가맹점을 아래 카테고리 중 하나로 분류하세요. 모르면 "기타".
카테고리: {cats}

가맹점:
{items}

JSON 객체 하나로만 답하세요: {{"가맹점명": "카테고리", ...}}"""


def classify_merchants(merchants: list[str]) -> dict[str, str]:
    if not merchants:
        return {}
    prompt = PROMPT.format(cats=", ".join(ALL_CATEGORIES), items="\n".join(f"- {m}" for m in merchants))
    try:
        if os.environ.get("ANTHROPIC_API_KEY"):
            raw = _anthropic(prompt)
        elif os.environ.get("OPENAI_API_KEY"):
            raw = _openai(prompt)
        else:
            return {}
        data = json.loads(raw[raw.find("{"): raw.rfind("}") + 1])
    except Exception as e:  # LLM 실패는 치명적이지 않다: '기타'로 남고 사람이 고친다
        print(f"[WARN] LLM 분류 실패: {e}")
        return {}
    return {m: c for m, c in data.items() if m in merchants and c in ALL_CATEGORIES}


def _anthropic(prompt: str) -> str:
    r = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={"x-api-key": os.environ["ANTHROPIC_API_KEY"], "anthropic-version": "2023-06-01",
                 "content-type": "application/json"},
        json={"model": os.environ.get("CARD_REPORT_MODEL", "claude-haiku-4-5"), "max_tokens": 2048,
              "messages": [{"role": "user", "content": prompt}]},
        timeout=60,
    )
    r.raise_for_status()
    return "".join(b.get("text", "") for b in r.json()["content"])


def _openai(prompt: str) -> str:
    r = requests.post(
        "https://api.openai.com/v1/chat/completions",
        headers={"Authorization": f"Bearer {os.environ['OPENAI_API_KEY']}"},
        json={"model": os.environ.get("CARD_REPORT_MODEL", "gpt-4o-mini"),
              "response_format": {"type": "json_object"},
              "messages": [{"role": "user", "content": prompt}]},
        timeout=60,
    )
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"]
