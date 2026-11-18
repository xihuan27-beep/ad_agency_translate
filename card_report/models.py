from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import date


@dataclass
class Transaction:
    date: date
    merchant: str
    amount: int                      # 이번 명세서 청구액(원). 취소/환불은 음수
    category: str = "미분류"
    installment_no: int | None = None   # 할부 회차 (3/6 의 3)
    installment_total: int | None = None  # 할부 총 개월 (3/6 의 6)
    original_amount: int | None = None   # 할부 원금(이용금액)
    time: str | None = None             # "HH:MM" — 대부분의 명세서엔 없음
    category_source: str = "rule"       # rule | llm | override | default

    @property
    def is_installment(self) -> bool:
        return bool(self.installment_total and self.installment_total > 1)

    @property
    def weekday(self) -> int:
        return self.date.weekday()  # 0=월

    def to_dict(self) -> dict:
        d = asdict(self)
        d["date"] = self.date.isoformat()
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "Transaction":
        d = dict(d)
        d["date"] = date.fromisoformat(d["date"])
        return cls(**d)


@dataclass
class Statement:
    transactions: list[Transaction]
    billing_month: str | None = None   # "2026-11" 결제월
    period_start: date | None = None   # 이용기간
    period_end: date | None = None
    stated_total: int | None = None    # 명세서에 적힌 청구 합계 (검산용)
    warnings: list[str] = field(default_factory=list)

    @property
    def total(self) -> int:
        return sum(t.amount for t in self.transactions)
