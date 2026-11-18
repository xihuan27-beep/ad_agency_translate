"""A → B → C → D 를 한 번에 돌리고 리포트(JSON + Markdown + 거래 CSV)를 만든다."""
from __future__ import annotations

import csv
import json
from pathlib import Path

from . import actions, forecast, parser, patterns
from .models import Statement


def load_overrides(path: str | Path | None) -> dict[str, str]:
    """사람이 고친 분류: CSV(merchant,category) 또는 이전 run의 transactions.csv."""
    if not path:
        return {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        return {r["merchant"].strip(): r["category"].strip()
                for r in csv.DictReader(f) if r.get("merchant") and r.get("category")}


def run(statement_path: str | Path, prev_report_path: str | Path | None = None,
        overrides_path: str | Path | None = None, use_llm: bool = False,
        default_year: int | None = None) -> tuple[dict, Statement]:
    prev = json.loads(Path(prev_report_path).read_text(encoding="utf-8")) if prev_report_path else None
    stmt = parser.load_statement(statement_path, default_year, load_overrides(overrides_path), use_llm)

    summary = patterns.summarize(stmt)
    findings = patterns.find_issues(stmt)
    fc = forecast.forecast_next(stmt, prev)
    acts = actions.pick_actions(findings)

    report = {
        "billing_month": stmt.billing_month,
        "period": [d.isoformat() if d else None for d in (stmt.period_start, stmt.period_end)],
        "summary": summary,
        "findings": [f.to_dict() for f in findings],
        "forecast": fc,
        "fixed_merchants": [f["merchant"] for f in fc["fixed"]],
        "actions": acts,
        "grading": actions.grade(prev["actions"], stmt) if prev else None,
        "forecast_accuracy": forecast.forecast_accuracy(prev, stmt) if prev else None,
        "previous_month": prev.get("billing_month") if prev else None,
        "needs_review": [t.to_dict() for t in stmt.transactions if t.category_source in ("default", "llm")],
        "warnings": stmt.warnings,
    }
    return report, stmt


def write_outputs(report: dict, stmt: Statement, out_dir: str | Path) -> dict[str, Path]:
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    paths = {"json": out / "report.json", "md": out / "report.md", "csv": out / "transactions.csv"}
    paths["json"].write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    paths["md"].write_text(render_markdown(report), encoding="utf-8")
    with open(paths["csv"], "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["date", "merchant", "amount", "installment", "category", "category_source"])
        for t in stmt.transactions:
            inst = f"{t.installment_no or 1}/{t.installment_total}" if t.is_installment else ""
            w.writerow([t.date.isoformat(), t.merchant, t.amount, inst, t.category, t.category_source])
    return paths


def render_markdown(r: dict) -> str:
    L = [f"# 카드 소비 리포트 — {r['billing_month']}", ""]
    if r["period"][0]:
        L += [f"이용기간 {r['period'][0]} ~ {r['period'][1]} · 청구 합계 **{r['summary']['total_billed']:,}원**", ""]

    if g := r.get("grading"):
        L += [f"## 지난달({r['previous_month']}) 제안 이행 점수: **{g['score']}**", ""]
        for x in g["results"]:
            mark = {"지킴": "✅", "못 지킴": "❌"}.get(x["status"], "➖")
            L.append(f"- {mark} {x['text']}  \n  {x['detail']}")
        L.append("")
    if fa := r.get("forecast_accuracy"):
        L += [f"지난달 예상 {fa['predicted']:,}원 → 실제 {fa['actual']:,}원 "
              f"(오차 {fa['error_pct']:+.1f}%, {'10% 이내' if fa['within_10pct'] else '10% 초과'})", ""]

    L += ["## 다음 달 실행 항목", ""]
    if not r["actions"]:
        L.append("- 이번 달은 줄일 만한 반복 지출을 찾지 못했습니다.")
    for i, a in enumerate(r["actions"], 1):
        L.append(f"{i}. **{a['text']}**  \n   예상 절감 {a['expected_saving']:,}원 · 근거 {a['basis']}")
    total_saving = sum(a["expected_saving"] for a in r["actions"])
    L += ["", f"세 가지를 모두 지키면 약 **{total_saving:,}원**을 덜 씁니다 (가정: 횟수 40% 감축 × 평균 단가).", ""]

    fc = r["forecast"]
    L += ["## 다음 명세서 예상 청구액", "",
          f"**{fc['next_total']:,}원** = 할부 {fc['installments_sum']:,} + 정기 {fc['fixed_sum']:,} "
          f"+ 변동 기준선 {fc['variable_baseline']:,}", f"<sub>{fc['method']}</sub>", ""]
    for i in fc["installments"]:
        L.append(f"- 할부 {i['merchant']} {i['next_no']}/{i['total']}회차 {i['amount']:,}원")
    for f in fc["fixed"]:
        L.append(f"- 정기 {f['merchant']} {f['amount']:,}원")
    L.append("")

    L += ["## 카테고리별 지출 (할부 승계분 제외)", "", "| 카테고리 | 건수 | 금액 |", "|---|---:|---:|"]
    for c in r["summary"]["categories"]:
        L.append(f"| {c['category']} | {c['count']} | {c['amount']:,} |")
    L += ["", "| " + " | ".join(r["summary"]["weekday_amounts"]) + " |",
          "|" + "---:|" * 7, "| " + " | ".join(f"{v:,}" for v in r["summary"]["weekday_amounts"].values()) + " |", ""]
    if not r["summary"]["has_time"]:
        L += ["<sub>이 명세서에는 결제 시각이 없어 시간대 분석은 하지 않았습니다.</sub>", ""]

    if r["findings"]:
        L += ["## 아쉬운 지점 전체", ""]
        for f in r["findings"]:
            L.append(f"- {f['title']} — 절감 가능 {f['saving']:,}원")
            L += [f"  - {e}" for e in f["evidence"]]
        L.append("")

    L += ["## 사람이 확인할 것", ""]
    if r["needs_review"]:
        L.append("규칙으로 분류하지 못한 거래입니다. `transactions.csv`의 category를 고친 뒤 `--overrides`로 다시 돌리세요.")
        L += [f"- {t['date']} {t['merchant']} {t['amount']:,}원 → {t['category']} ({t['category_source']})"
              for t in r["needs_review"]]
    else:
        L.append("- 분류 확인이 필요한 거래 없음")
    L.append("- 실행 항목이 비현실적이면 `report.json`의 해당 항목 `approved`를 `false`로 바꾸세요. 다음 달 채점에서 빠집니다.")
    for w in r["warnings"]:
        L.append(f"- ⚠️ {w}")
    return "\n".join(L) + "\n"
