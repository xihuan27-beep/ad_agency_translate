"""사용법:
  python -m card_report 2026-10.pdf -o out/2026-10
  python -m card_report 2026-11.pdf --prev out/2026-10/report.json -o out/2026-11
  python -m card_report 2026-11.pdf --prev ... --overrides out/2026-11/transactions.csv -o out/2026-11
"""
import argparse
import sys

from .pipeline import run, write_outputs


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="card_report", description="카드 명세서 → 다음 달 실행 항목 + 지난달 채점")
    ap.add_argument("statement", help="카드 명세서 PDF 또는 CSV")
    ap.add_argument("--prev", help="지난달 report.json (있으면 이행 점수를 매김)")
    ap.add_argument("--overrides", help="사람이 고친 분류 CSV (merchant,category)")
    ap.add_argument("--llm", action="store_true", help="규칙으로 분류 못 한 가맹점을 LLM에 물어봄")
    ap.add_argument("--year", type=int, help="날짜에 연도가 없을 때 쓸 연도")
    ap.add_argument("-o", "--out", default="card_report_out")
    a = ap.parse_args(argv)

    report, stmt = run(a.statement, a.prev, a.overrides, a.llm, a.year)
    paths = write_outputs(report, stmt, a.out)
    print(paths["md"].read_text(encoding="utf-8"))
    print(f"\n저장: {', '.join(str(p) for p in paths.values())}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
