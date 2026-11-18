"""A. 명세서 파서 — PDF/CSV에서 거래를 추출하고 카테고리를 붙인다.

카드사마다 PDF 레이아웃이 다르므로 '날짜로 시작하고 금액으로 끝나는 줄'이라는
최소 공통 규칙으로 추출한다. 못 읽은 줄은 경고로 남겨 사람이 확인하게 한다.
"""
from __future__ import annotations

import csv
import re
from datetime import date
from pathlib import Path

from .categories import classify_rule
from .models import Statement, Transaction

DATE_RE = re.compile(
    r"^(?:(?P<y>\d{4}|\d{2})[./-])?(?P<m>\d{1,2})[./-](?P<d>\d{1,2})(?=\s)"
)
TIME_RE = re.compile(r"^\s*(?P<t>[0-2]\d:[0-5]\d)(?::[0-5]\d)?\s")
AMOUNT_RE = re.compile(r"^-?\d{1,3}(?:,\d{3})+$|^-?\d+$")
INSTALL_RE = re.compile(r"(?:할부\s*)?(?P<no>\d{1,2})\s*/\s*(?P<tot>\d{1,2})(?:\s*개월)?$")
INSTALL_TOTAL_ONLY_RE = re.compile(r"^(?P<tot>\d{1,2})개월$")

BILLING_RE = re.compile(r"(?P<y>20\d{2})\s*년\s*(?P<m>\d{1,2})\s*월\s*(?:분|청구|결제|이용대금)")
PERIOD_RE = re.compile(
    r"이용\s*기간\s*:?\s*(?P<a>20\d{2}[./-]\d{1,2}[./-]\d{1,2})\s*[~\-]\s*(?P<b>20\d{2}[./-]\d{1,2}[./-]\d{1,2})"
)
TOTAL_RE = re.compile(r"(?:청구\s*금액|결제\s*하실\s*금액|이번\s*달\s*결제\s*금액|합\s*계)\s*:?\s*(?P<v>-?[\d,]+)\s*원?")
SKIP_WORDS = ("합계", "소계", "총액", "청구금액", "결제하실", "이용기간", "결제일", "페이지")


def _to_int(s: str) -> int:
    return int(s.replace(",", ""))


def _to_date(s: str) -> date:
    y, m, d = re.split(r"[./-]", s)
    return date(int(y), int(m), int(d))


def parse_text(text: str, default_year: int | None = None) -> Statement:
    stmt = Statement(transactions=[])
    if m := BILLING_RE.search(text):
        stmt.billing_month = f"{int(m['y']):04d}-{int(m['m']):02d}"
    if m := PERIOD_RE.search(text):
        stmt.period_start, stmt.period_end = _to_date(m["a"]), _to_date(m["b"])
    if m := TOTAL_RE.search(text):
        stmt.stated_total = _to_int(m["v"])

    year_hint = default_year or (stmt.period_end.year if stmt.period_end else None) or date.today().year

    for raw in text.splitlines():
        line = raw.strip()
        if not line or any(w in line for w in SKIP_WORDS):
            continue
        dm = DATE_RE.match(line)
        if not dm:
            continue
        tx = _parse_line(line, dm, year_hint, stmt)
        if tx is None:
            stmt.warnings.append(f"읽지 못한 줄: {line}")
        else:
            stmt.transactions.append(tx)
    return stmt


def _parse_line(line: str, dm: re.Match, year_hint: int, stmt: Statement) -> Transaction | None:
    y = dm["y"]
    year = (2000 + int(y) if len(y) == 2 else int(y)) if y else year_hint
    try:
        d = date(year, int(dm["m"]), int(dm["d"]))
    except ValueError:
        return None
    # 연도가 없는 MM.DD 형식이고 이용기간을 알면, 기간 밖이면 전년도로 본다 (12월→1월 경계)
    if not y and stmt.period_end and d > stmt.period_end:
        d = d.replace(year=d.year - 1)

    rest = line[dm.end():]
    t = None
    if tm := TIME_RE.match(rest):
        t, rest = tm["t"], rest[tm.end():]

    tokens = rest.split()
    amounts: list[int] = []
    inst_no = inst_tot = None
    # 뒤에서부터 숫자/할부 토큰을 걷어낸다
    while tokens:
        tok = tokens[-1]
        if AMOUNT_RE.match(tok):
            amounts.insert(0, _to_int(tok))
        elif (im := INSTALL_RE.match(tok)) and inst_tot is None:
            inst_no, inst_tot = int(im["no"]), int(im["tot"])
        elif (im := INSTALL_TOTAL_ONLY_RE.match(tok)) and inst_tot is None:
            inst_tot = int(im["tot"])
        elif tok in ("일시불", "원", "할부"):
            pass
        else:
            break
        tokens.pop()
    merchant = " ".join(tokens).strip()
    if not merchant or not amounts:
        return None

    cancel = "취소" in merchant or "환불" in merchant
    billed = amounts[-1]
    original = amounts[0] if len(amounts) >= 2 else None
    # 할부 회차가 남은 숫자 열로 들어온 경우: [이용금액, 회차, 청구액] 처럼 가운데 작은 수
    if inst_tot and inst_no is None and len(amounts) >= 3 and amounts[1] <= inst_tot:
        inst_no = amounts[1]
    if cancel and billed > 0:
        billed = -billed
    return Transaction(
        date=d, merchant=merchant, amount=billed, time=t,
        installment_no=inst_no, installment_total=inst_tot,
        original_amount=original if (inst_tot and inst_tot > 1) else None,
    )


def read_pdf_text(path: Path) -> str:
    import pdfplumber

    with pdfplumber.open(str(path)) as pdf:
        return "\n".join((p.extract_text() or "") for p in pdf.pages)


def parse_csv(path: Path) -> Statement:
    """카드사 엑셀을 CSV로 저장한 파일. 헤더: date, merchant, amount [, time, installment, original_amount, category]"""
    stmt = Statement(transactions=[])
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            inst_no = inst_tot = None
            if (inst := (row.get("installment") or "").strip()) and (im := INSTALL_RE.match(inst)):
                inst_no, inst_tot = int(im["no"]), int(im["tot"])
            stmt.transactions.append(Transaction(
                date=_to_date(row["date"].strip()),
                merchant=row["merchant"].strip(),
                amount=_to_int(row["amount"].strip()),
                time=(row.get("time") or "").strip() or None,
                installment_no=inst_no, installment_total=inst_tot,
                original_amount=_to_int(row["original_amount"]) if row.get("original_amount") else None,
                category=(row.get("category") or "").strip() or "미분류",
                category_source="override" if (row.get("category") or "").strip() else "rule",
            ))
    return stmt


def classify(stmt: Statement, overrides: dict[str, str] | None = None, use_llm: bool = False) -> None:
    """카테고리 부여. 우선순위: 사람이 고친 값(override) > 규칙 > LLM > 기타."""
    overrides = overrides or {}
    unknown: list[Transaction] = []
    for tx in stmt.transactions:
        if tx.category_source == "override" and tx.category != "미분류":
            continue
        if tx.merchant in overrides:
            tx.category, tx.category_source = overrides[tx.merchant], "override"
        elif cat := classify_rule(tx.merchant):
            tx.category, tx.category_source = cat, "rule"
        else:
            unknown.append(tx)
    if unknown and use_llm:
        from .llm import classify_merchants

        mapping = classify_merchants(sorted({t.merchant for t in unknown}))
        for tx in unknown:
            if cat := mapping.get(tx.merchant):
                tx.category, tx.category_source = cat, "llm"
    for tx in unknown:
        if tx.category == "미분류":
            tx.category, tx.category_source = "기타", "default"


def load_statement(path: str | Path, default_year: int | None = None,
                   overrides: dict[str, str] | None = None, use_llm: bool = False) -> Statement:
    path = Path(path)
    if path.suffix.lower() == ".csv":
        stmt = parse_csv(path)
    elif path.suffix.lower() == ".pdf":
        stmt = parse_text(read_pdf_text(path), default_year)
        if not stmt.transactions:
            stmt.warnings.append("거래를 하나도 찾지 못했습니다. 이미지(스캔) PDF이거나 지원하지 않는 레이아웃일 수 있습니다. "
                                 "카드사 홈페이지의 엑셀 내려받기 → CSV 저장 후 넣어 주세요.")
    else:
        raise ValueError(f"지원하지 않는 형식: {path.suffix} (pdf, csv만 가능)")

    if stmt.transactions and not stmt.period_start:
        stmt.period_start = min(t.date for t in stmt.transactions)
        stmt.period_end = max(t.date for t in stmt.transactions)
    if not stmt.billing_month and stmt.period_end:
        stmt.billing_month = stmt.period_end.strftime("%Y-%m")
    if stmt.stated_total is not None and stmt.stated_total != stmt.total:
        stmt.warnings.append(
            f"검산 불일치: 명세서 합계 {stmt.stated_total:,}원 ≠ 추출 합계 {stmt.total:,}원 "
            f"(차이 {stmt.stated_total - stmt.total:,}원). 빠진 거래가 있는지 확인하세요."
        )
    classify(stmt, overrides, use_llm)
    return stmt
