import json
from datetime import date
from pathlib import Path

import pytest

from card_report import actions, forecast, parser, patterns
from card_report.models import Statement, Transaction
from card_report.pipeline import run, write_outputs

SAMPLES = Path(__file__).parents[1] / "samples"


def test_parse_line_variants():
    text = """가상카드 2026년 11월 청구
이용기간 2026.10.01 ~ 2026.10.31
2026.10.03 배달의민족 23,000
10.04 22:41 GS25 역삼점 4,800
2026.09.02 삼성전자 디지털프라자 1,200,000 2/6 200,000
2026.10.21 쿠팡 취소 -35,900
2026.10.22 쿠팡 취소 35,900
합계 1,000
"""
    s = parser.parse_text(text)
    assert s.billing_month == "2026-11"
    assert [t.amount for t in s.transactions] == [23000, 4800, 200000, -35900, -35900]
    assert s.transactions[1].time == "22:41" and s.transactions[1].date == date(2026, 10, 4)
    inst = s.transactions[2]
    assert (inst.installment_no, inst.installment_total, inst.original_amount) == (2, 6, 1200000)


def test_year_rollover_for_mmdd():
    s = parser.parse_text("이용기간 2026.12.15 ~ 2027.01.14\n12.20 스타벅스 5,000\n01.03 스타벅스 5,000")
    assert [t.date for t in s.transactions] == [date(2026, 12, 20), date(2027, 1, 3)]


def _tx(d, m, a, cat=None, **kw):
    t = Transaction(date=d, merchant=m, amount=a, **kw)
    t.category = cat or "기타"
    return t


def test_installment_carryover_excluded_from_patterns_but_in_forecast():
    s = Statement([_tx(date(2026, 9, 2), "가전", 200000, installment_no=2, installment_total=6),
                   _tx(date(2026, 10, 1), "카페", 5000, "카페")])
    assert [t.merchant for t in patterns.spending_tx(s)] == ["카페"]
    fc = forecast.forecast_next(s)
    assert fc["installments_sum"] == 200000 and fc["variable_baseline"] == 5000


def test_last_installment_not_forecast():
    s = Statement([_tx(date(2026, 5, 2), "가전", 100000, installment_no=6, installment_total=6)])
    assert forecast.forecast_next(s)["installments"] == []


def test_grading_count_and_absent():
    prev = [
        {"id": "a", "text": "배달", "rule": {"metric": "count", "category": "배달", "max": 2, "baseline": 5}},
        {"id": "b", "text": "넷플", "rule": {"metric": "absent", "merchant": "넷플릭스"}},
        {"id": "c", "text": "심야", "rule": {"metric": "count", "late_night": True, "max": 1, "baseline": 4}},
        {"id": "d", "text": "skip", "approved": False, "rule": {"metric": "absent", "merchant": "x"}},
    ]
    s = Statement([_tx(date(2026, 11, 1), "배민", 20000, "배달"),
                   _tx(date(2026, 11, 2), "배민", 20000, "배달"),
                   _tx(date(2026, 11, 3), "넷플릭스", 17000, "구독")])
    g = actions.grade(prev, s)
    status = {r["id"]: r["status"] for r in g["results"]}
    assert status == {"a": "지킴", "b": "못 지킴", "c": "판정 불가"}
    assert g["score"] == "1/2"


def test_actions_are_distinct_categories():
    days = [date(2026, 10, d) for d in range(1, 29)]
    s = Statement([_tx(d, "배민", 20000, "배달") for d in days if d.weekday() >= 5]
                  + [_tx(d, "스벅", 5000, "카페") for d in days])
    acts = actions.pick_actions(patterns.find_issues(s))
    cats = [a["rule"].get("category") for a in acts]
    assert len(cats) == len(set(cats))


@pytest.mark.skipif(not (SAMPLES / "statement_2026-10.pdf").exists(), reason="samples not generated")
def test_two_month_end_to_end(tmp_path):
    r1, s1 = run(SAMPLES / "statement_2026-10.pdf")
    assert not r1["warnings"], r1["warnings"]  # 검산 일치
    assert len(r1["actions"]) == 3
    p = write_outputs(r1, s1, tmp_path / "10")
    r2, _ = run(SAMPLES / "statement_2026-11.pdf", prev_report_path=p["json"])
    assert r2["grading"]["judged"] == 3
    assert r2["grading"]["kept"] >= 1          # 피치 목표: 3건 중 1건 이상
    assert r2["forecast_accuracy"]["actual"] == r2["summary"]["total_billed"]
    json.dumps(r2, ensure_ascii=False)
