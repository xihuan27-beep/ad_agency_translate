"""D. 실행·점검 — 실행 항목 3개를 고르고, 지난달 실행 항목을 이번 명세서로 채점한다.

실행 항목은 반드시 '다음 명세서만으로 판정 가능한' 규칙이어야 한다.
("장보기 목록 쓰기" 같은 항목은 명세서로 채점할 수 없으므로 만들지 않는다.)
"""
from __future__ import annotations

from .models import Statement
from .patterns import Finding, WEEKDAYS, LATE_NIGHT, spending_tx

N_ACTIONS = 3


def pick_actions(findings: list[Finding], n: int = N_ACTIONS) -> list[dict]:
    """절감액 큰 순으로 고르되, 같은 카테고리를 두 번 고르지 않는다."""
    chosen, used_cats = [], set()
    for f in findings:
        key = f.category or f.kind
        if key in used_cats:
            continue
        chosen.append({
            "id": f.id, "text": f.title, "expected_saving": f.saving,
            "rule": f.target, "approved": True,
            "basis": f"{f.tx_count}건 · {f.amount:,}원",
        })
        used_cats.add(key)
        if len(chosen) == n:
            break
    return chosen


def _measure(rule: dict, stmt: Statement) -> int:
    txs = spending_tx(stmt)
    if rule["metric"] == "absent":
        return sum(1 for t in stmt.transactions if t.merchant == rule["merchant"] and t.amount > 0)
    sel = txs
    if cat := rule.get("category"):
        sel = [t for t in sel if t.category == cat]
    if days := rule.get("weekdays"):
        sel = [t for t in sel if t.weekday in days]
    if rule.get("late_night"):
        sel = [t for t in sel if t.time and (int(t.time[:2]) >= LATE_NIGHT[0] or int(t.time[:2]) < LATE_NIGHT[1])]
    return len(sel)


def grade(prev_actions: list[dict], stmt: Statement) -> dict:
    results = []
    for a in prev_actions:
        if not a.get("approved", True):
            continue
        rule = a["rule"]
        if rule.get("late_night") and not any(t.time for t in stmt.transactions):
            results.append({**_head(a), "status": "판정 불가", "detail": "이번 명세서에 결제 시각이 없음"})
            continue
        actual = _measure(rule, stmt)
        if rule["metric"] == "absent":
            ok = actual == 0
            detail = "결제 없음 (해지 확인)" if ok else f"여전히 {actual}건 결제됨"
        else:
            ok = actual <= rule["max"]
            where = rule.get("category") or "심야"
            if days := rule.get("weekdays"):
                where = "/".join(WEEKDAYS[d] for d in days) + " " + where
            detail = f"{where} {rule['baseline']}회 → {actual}회 (목표 {rule['max']}회 이하)"
        results.append({**_head(a), "status": "지킴" if ok else "못 지킴", "detail": detail})
    judged = [r for r in results if r["status"] != "판정 불가"]
    kept = sum(r["status"] == "지킴" for r in judged)
    return {"results": results, "kept": kept, "judged": len(judged),
            "score": f"{kept}/{len(judged)}" if judged else "-"}


def _head(a: dict) -> dict:
    return {"id": a["id"], "text": a["text"]}
