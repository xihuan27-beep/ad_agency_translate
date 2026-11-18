"""C. 지출 예측 — 다음 명세서(다음 결제월) 예상 청구액.

예상액 = 남은 할부의 다음 회차 + 정기 결제 + 변동 지출 기준선
변동 지출 기준선은 이번 달 값(지난달 리포트가 있으면 두 달 평균)을 쓴다.
"""
from __future__ import annotations

from .models import Statement

FIXED_CATEGORIES = {"구독", "통신", "보험"}


def forecast_next(stmt: Statement, prev: dict | None = None) -> dict:
    installments, fixed, variable = [], [], 0
    prev_merchants = set((prev or {}).get("fixed_merchants", []))
    for t in stmt.transactions:
        if t.is_installment:
            if (t.installment_no or 1) < (t.installment_total or 1):
                installments.append({
                    "merchant": t.merchant, "next_no": (t.installment_no or 1) + 1,
                    "total": t.installment_total, "amount": t.amount,
                })
        elif t.amount > 0 and (t.category in FIXED_CATEGORIES or t.merchant in prev_merchants):
            fixed.append({"merchant": t.merchant, "amount": t.amount})
        else:
            variable += t.amount  # 취소(음수)도 반영

    prev_var = (prev or {}).get("forecast", {}).get("variable_this_month")
    baseline = round((variable + prev_var) / 2) if prev_var is not None else variable
    inst_sum = sum(i["amount"] for i in installments)
    fixed_sum = sum(f["amount"] for f in fixed)
    return {
        "next_total": inst_sum + fixed_sum + baseline,
        "installments": installments,
        "installments_sum": inst_sum,
        "fixed": fixed,
        "fixed_sum": fixed_sum,
        "variable_baseline": baseline,
        "variable_this_month": variable,
        "method": "할부 다음 회차 + 정기 결제 + 변동 지출 기준선" + (" (2개월 평균)" if prev_var is not None else " (이번 달 값)"),
    }


def forecast_accuracy(prev_report: dict, stmt: Statement) -> dict | None:
    predicted = prev_report.get("forecast", {}).get("next_total")
    if not predicted:
        return None
    actual = stmt.total
    err = (actual - predicted) / actual if actual else 0.0
    return {"predicted": predicted, "actual": actual, "error_pct": round(err * 100, 1),
            "within_10pct": abs(err) <= 0.10}
