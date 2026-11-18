"""B. 패턴 분석 — 카테고리·요일·시간대로 묶고, '아쉬운 지점'과 절감액을 근거와 함께 낸다.

모든 절감액은 '횟수를 줄이면 평균 단가만큼 준다'는 단순 가정으로 계산하며,
계산에 쓴 거래 건수와 금액을 함께 남긴다 (피치 4쪽: 확인할 근거).
"""
from __future__ import annotations

import math
from collections import defaultdict
from dataclasses import asdict, dataclass, field

from .categories import DISCRETIONARY
from .models import Statement, Transaction

WEEKDAYS = ["월", "화", "수", "목", "금", "토", "일"]
REDUCTION = 0.4          # 횟수 40% 감축을 가정
MIN_COUNT = 3            # 이 횟수 미만이면 '습관'으로 보지 않음
LATE_NIGHT = (22, 4)     # 22:00 ~ 03:59


@dataclass
class Finding:
    id: str
    kind: str                 # category_count | weekday_category | late_night | subscription
    title: str
    category: str | None
    tx_count: int
    amount: int
    saving: int
    target: dict              # D 단계에서 채점할 수 있는 측정 규칙
    evidence: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)


def spending_tx(stmt: Statement) -> list[Transaction]:
    """패턴 분석 대상: 할부 승계분·취소 제외, 이번 달 새로 쓴 돈."""
    return [t for t in stmt.transactions if t.amount > 0 and not (t.is_installment and (t.installment_no or 1) > 1)]


def summarize(stmt: Statement) -> dict:
    by_cat: dict[str, list[int]] = defaultdict(list)
    by_wd = [0] * 7
    for t in spending_tx(stmt):
        by_cat[t.category].append(t.amount)
        by_wd[t.weekday] += t.amount
    cats = sorted(({"category": c, "count": len(v), "amount": sum(v)} for c, v in by_cat.items()),
                  key=lambda x: -x["amount"])
    has_time = any(t.time for t in stmt.transactions)
    return {
        "total_billed": stmt.total,
        "categories": cats,
        "weekday_amounts": dict(zip(WEEKDAYS, by_wd)),
        "has_time": has_time,
    }


def _ev(txs: list[Transaction], limit: int = 6) -> list[str]:
    out = [f"{t.date:%m/%d}({WEEKDAYS[t.weekday]}) {t.merchant} {t.amount:,}원" for t in txs[:limit]]
    if len(txs) > limit:
        out.append(f"… 외 {len(txs) - limit}건")
    return out


def _cut(count: int) -> int:
    return max(1, round(count * REDUCTION))


def find_issues(stmt: Statement) -> list[Finding]:
    txs = spending_tx(stmt)
    findings: list[Finding] = []
    by_cat: dict[str, list[Transaction]] = defaultdict(list)
    for t in txs:
        by_cat[t.category].append(t)

    # 1) 재량 카테고리의 잦은 결제
    for cat, items in by_cat.items():
        if cat not in DISCRETIONARY or cat == "구독" or len(items) < MIN_COUNT:
            continue
        n, amt = len(items), sum(t.amount for t in items)
        cut = _cut(n)
        findings.append(Finding(
            id=f"cat:{cat}", kind="category_count",
            title=f"{cat} {n}회 · {amt:,}원 — 횟수를 {n - cut}회 이하로",
            category=cat, tx_count=n, amount=amt, saving=math.floor(amt / n * cut),
            target={"metric": "count", "category": cat, "max": n - cut, "baseline": n},
            evidence=_ev(sorted(items, key=lambda t: t.date)),
        ))

    # 2) 특정 요일 쏠림: 재량 카테고리 결제 중 한 요일군에 절반 이상
    groups = {"주말": {5, 6}, "금요일": {4}, "평일": {0, 1, 2, 3, 4}}
    for cat, items in by_cat.items():
        if cat not in DISCRETIONARY or cat == "구독" or len(items) < MIN_COUNT + 1:
            continue
        for gname, days in groups.items():
            if gname == "평일":
                continue  # 평일 쏠림은 정상 분포라 의미가 약하다
            sel = [t for t in items if t.weekday in days]
            expected_share = len(days) / 7
            share = len(sel) / len(items)
            # share가 90% 이상이면 '카테고리 전체' 항목과 같은 말이라 따로 내지 않는다
            if len(sel) >= MIN_COUNT and max(0.5, expected_share * 1.75) <= share < 0.9:
                amt = sum(t.amount for t in sel)
                cut = _cut(len(sel))
                findings.append(Finding(
                    id=f"wd:{cat}:{gname}", kind="weekday_category",
                    title=f"{gname} {cat} {len(sel)}회 ({share:.0%}) — {gname}엔 {len(sel) - cut}회 이하로",
                    category=cat, tx_count=len(sel), amount=amt, saving=math.floor(amt / len(sel) * cut),
                    target={"metric": "count", "category": cat, "weekdays": sorted(days),
                            "max": len(sel) - cut, "baseline": len(sel)},
                    evidence=_ev(sel),
                ))

    # 3) 심야 결제 (시간 정보가 있을 때만)
    late = [t for t in txs if t.time and (int(t.time[:2]) >= LATE_NIGHT[0] or int(t.time[:2]) < LATE_NIGHT[1])
            and t.category in DISCRETIONARY]
    if len(late) >= MIN_COUNT:
        amt = sum(t.amount for t in late)
        cut = _cut(len(late))
        findings.append(Finding(
            id="late_night", kind="late_night",
            title=f"밤 10시 이후 결제 {len(late)}회 · {amt:,}원 — {len(late) - cut}회 이하로",
            category=None, tx_count=len(late), amount=amt, saving=math.floor(amt / len(late) * cut),
            target={"metric": "count", "late_night": True, "max": len(late) - cut, "baseline": len(late)},
            evidence=_ev(late),
        ))

    # 4) 구독: 하나씩 '해지 후보'
    for t in by_cat.get("구독", []):
        findings.append(Finding(
            id=f"sub:{t.merchant}", kind="subscription",
            title=f"구독 {t.merchant} {t.amount:,}원 — 쓰지 않으면 해지",
            category="구독", tx_count=1, amount=t.amount, saving=t.amount,
            target={"metric": "absent", "merchant": t.merchant},
            evidence=_ev([t]),
        ))

    findings.sort(key=lambda f: -f.saving)
    return findings
