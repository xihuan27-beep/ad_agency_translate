# 카드 소비 리포트 에이전트

카드 명세서를 넣으면 **다음 달 실행 항목 3가지**를 제안하고, 다음 달 명세서로 **지켰는지 채점**합니다.

```
명세서 PDF/CSV + 지난달 report.json
  → A 파서(parser.py)       거래 추출 · 카테고리 분류 · 합계 검산
  → B 패턴(patterns.py)     카테고리/요일/시간대 묶기 · 아쉬운 지점 · 절감액(근거 거래 포함)
  → C 예측(forecast.py)     할부 다음 회차 + 정기 결제 + 변동 기준선 = 다음 명세서 예상액
  → D 실행·점검(actions.py) 실행 항목 3개(명세서로 판정 가능한 규칙) · 지난달 항목 채점
  → report.md / report.json / transactions.csv
```

## 실행

```bash
pip install pdfplumber reportlab requests
python card_report/samples/make_samples.py          # 가상 명세서 2개월치 생성
python -m card_report card_report/samples/statement_2026-10.pdf -o out/10
python -m card_report card_report/samples/statement_2026-11.pdf --prev out/10/report.json -o out/11
python -m pytest card_report/tests
```

## 사람의 확인·승인

- **분류 오류**: `transactions.csv`의 `category`를 고친 뒤 `--overrides out/11/transactions.csv`로 다시 실행.
- **실행 항목 현실성**: `report.json`의 `actions[].approved`를 `false`로 바꾸면 다음 달 채점에서 제외.
- **검산 불일치**: 명세서 합계와 추출 합계가 다르면 경고가 뜹니다. 빠진 줄은 "읽지 못한 줄"로 표시됩니다.

## 설계상의 결정과 한계

- 계산(합계·절감액·예측·채점)은 모두 결정론적 코드입니다. LLM은 `--llm`을 줄 때만, 규칙으로 분류 못 한
  가맹점명만 보냅니다(금액·날짜는 보내지 않음). `ANTHROPIC_API_KEY` 또는 `OPENAI_API_KEY` 필요.
- PDF 파서는 "날짜로 시작하고 금액으로 끝나는 줄" 규칙입니다. 카드사별 실제 양식으로는 아직 검증하지 않았고,
  스캔(이미지) PDF·암호 PDF는 읽지 못합니다. 이때는 카드사 엑셀을 CSV(`date,merchant,amount[,time,installment]`)로 넣으세요.
- 국내 명세서 대부분은 **결제 시각이 없어** 시간대 분석·심야 항목은 시각이 있는 입력에서만 동작합니다.
- 절감액 가정: 해당 거래 횟수 40% 감축 × 평균 단가. 구독은 해지 시 월 요금 전액.
- 예측은 "아무것도 바꾸지 않았을 때"의 기준선입니다. 실행 항목을 지키면 실제 청구액은 예측보다 낮아지는 게 정상입니다.
