"""테스트용 가상 명세서 PDF 두 달 치를 만든다 (실제 카드사 양식 아님).

  python card_report/samples/make_samples.py
"""
from datetime import date, timedelta
from pathlib import Path

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas

FONT = "HYSMyeongJo-Medium"
HERE = Path(__file__).parent


def month_rows(year, month, weekend_delivery, weekday_cafe, subs, extra):
    start = date(year, month, 1)
    rows = []
    d = start
    i = 0
    while d.month == month:
        wd = d.weekday()
        if wd >= 5 and weekend_delivery and i % 1 == 0:
            rows.append((d, "배달의민족", 23000 + (d.day % 3) * 2500))
        if wd < 5 and d.day % weekday_cafe == 0:
            rows.append((d, "스타벅스 강남역점", 6100))
        if wd == 2:
            rows.append((d, "GS25 역삼점", 4800))
        if d.day in (5, 19):
            rows.append((d, "이마트 성수점", 68000))
        if d.day == 12:
            rows.append((d, "카카오T 택시", 15800))
        d += timedelta(days=1)
        i += 1
    rows += [(date(year, month, 3), m, a) for m, a in subs]
    rows += extra
    return sorted(rows, key=lambda r: r[0])


def draw(path, title, period, rows, installments):
    pdfmetrics.registerFont(UnicodeCIDFont(FONT))
    c = canvas.Canvas(str(path))
    total = sum(a for _, _, a in rows) + sum(r[-1] for r in installments)
    y = 800

    def line(s, size=9):
        nonlocal y
        if y < 50:
            c.showPage(); y = 800
        c.setFont(FONT, size); c.drawString(40, y, s); y -= 14

    line(title, 13); line(period); line(f"청구금액 {total:,}원"); y -= 8
    line("이용일자   이용가맹점                     이용금액    할부   청구금액")
    for d, m, a in rows:
        line(f"{d:%Y.%m.%d}  {m}  {a:,}")
    for d, m, orig, inst, billed in installments:
        line(f"{d:%Y.%m.%d}  {m}  {orig:,}  {inst}  {billed:,}")
    c.save()


if __name__ == "__main__":
    oct_rows = month_rows(2026, 10, True, 1, [("넷플릭스", 17000), ("유튜브프리미엄", 14900)],
                          [(date(2026, 10, 9), "을지로 노가리집", 42000),
                           (date(2026, 10, 20), "쿠팡", 35900),
                           (date(2026, 10, 21), "쿠팡 취소", -35900)])
    oct_inst = [(date(2026, 9, 2), "삼성전자 디지털프라자", 1200000, "2/6", 200000)]
    draw(HERE / "statement_2026-10.pdf", "가상카드 2026년 11월 청구 명세서",
         "이용기간 2026.10.01 ~ 2026.10.31", oct_rows, oct_inst)

    nov_rows = [r for r in month_rows(2026, 11, True, 2, [("유튜브프리미엄", 14900)],
                                      [(date(2026, 11, 14), "을지로 노가리집", 38000)])
                if not (r[1] == "배달의민족" and r[0].day % 3 == 0)]
    nov_inst = [(date(2026, 9, 2), "삼성전자 디지털프라자", 1200000, "3/6", 200000)]
    draw(HERE / "statement_2026-11.pdf", "가상카드 2026년 12월 청구 명세서",
         "이용기간 2026.11.01 ~ 2026.11.30", nov_rows, nov_inst)
    print("created", *HERE.glob("statement_*.pdf"))
