"""리뷰 CSV에서 '깊이' 계열 표현을 3단계로 집계하고 맥락을 분류한다.

입력: collect_serf_reviews.py가 만든 CSV (review_id,date,rating,treatment,hospital,body)
출력: 표준출력(마크다운 표) + --json 경로에 기계용 결과

맥락 분류는 규칙 기반 1차 분류다. 결과의 `needs_review=True` 행은 사람이 앞뒤 문장을 읽고
최종 판정해야 한다(규칙 기준은 reviews/serf_depth_analysis.md 참고).
"""
import argparse, csv, json, re, statistics, sys
from collections import Counter, defaultdict

TIERS = {
    # (a) 정확히 '깊이'. '깊이감'은 (b)로 넘긴다.
    "a_exact": re.compile(r"깊이(?!감)"),
    # (b) 같은 어근 변형(지정된 5개만)
    "b_variant": re.compile(r"깊게|깊은|깊숙이|깊이감|깊어"),
    # (c) 반대 표현
    "c_opposite": re.compile(r"얕게|얕은"),
}
# 지정 목록 밖의 '깊-' 활용형(깊고, 깊다, 깊지, 깊었, 깊네 등). 집계 누락 진단용.
OTHER_DEEP = re.compile(r"깊(?!이|게|은|숙이|어)[가-힣]")

# 우선순위 순서대로 처음 맞는 규칙을 채택한다.
RULES = [
    ("idiom", "시술과 무관한 관용 표현",
     re.compile(r"깊[이은]\s?(감사|반성|고민|생각|공감|신뢰|인상|사과|존경)|깊은\s?(밤|잠|뜻|관심|애정)")),
    ("pain", "통증과 연결",
     re.compile(r"아프|아팠|아파|통증|따갑|따끔|욱신|찌릿|뜨거|참기|고통|마취")),
    ("staff", "의료진 설명/조절",
     re.compile(r"조절|맞춰|설명|원장|선생님|쌤|실장|상담|조정|세팅|레벨")),
    ("concern", "깊은 주름이나 처짐(시술 전 고민)",
     re.compile(r"주름|처짐|처져|쳐진|팔자|고랑|심술보|불독|이중턱|볼살|패인|파인")),
    ("effect", "시술/에너지가 깊이 들어간다(효과 체감)",
     re.compile(r"들어가|들어오|들어온|전달|닿|타겟|근막|SMAS|스마스|진피|층|효과|느낌|당겨|리프팅|탄력|올라")),
]
LABELS = {k: v for k, _, v in [(r[0], None, r[1]) for r in RULES]}
LABELS["other"] = "기타"

POS = re.compile(r"만족|좋|추천|효과\s?(있|봤|짱)|재방문|또\s?받|최고|친절|자연스럽")
NEG = re.compile(r"불만|별로|후회|효과\s?(없|못)|실망|비추|아깝|돈\s?낭비|부작용|붓기\s?오래")

SENT_SPLIT = re.compile(r"(?<=[.!?~。])\s+|\n+")

def sentences(text):
    return [s.strip() for s in SENT_SPLIT.split(text or "") if s.strip()]

def classify(window, hit_sentence):
    # 관용 표현은 등장 문장 자체만 본다(앞뒤 문장의 통증 단어에 끌려가지 않도록).
    if RULES[0][2].search(hit_sentence):
        return "idiom", False
    for key, _, rx in RULES[1:]:
        if rx.search(hit_sentence):
            return key, False
    for key, _, rx in RULES[1:]:
        if rx.search(window):
            return key, True  # 앞뒤 문장에서만 근거가 나옴 -> 사람 검토 권장
    return "other", True

def polarity(rating, text):
    try:
        r = float(rating)
        if r >= 4: return "pos"
        if r <= 2: return "neg"
    except (TypeError, ValueError):
        pass
    p, n = bool(POS.search(text)), bool(NEG.search(text))
    return "pos" if p and not n else "neg" if n and not p else "neutral"

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("csv", nargs="?", default="reviews/serf_reviews.csv")
    ap.add_argument("--json", default=None)
    a = ap.parse_args()
    rows = list(csv.DictReader(open(a.csv, encoding="utf-8-sig")))
    n = len(rows)

    tier = {k: {"mentions": 0, "reviews": 0} for k in TIERS}
    other_forms = Counter()
    hits = []
    for row in rows:
        body = row.get("body", "")
        for k, rx in TIERS.items():
            c = len(rx.findall(body))
            tier[k]["mentions"] += c
            tier[k]["reviews"] += c > 0
        other_forms.update(m.group(0) for m in OTHER_DEEP.finditer(body))
        sents = sentences(body)
        for i, s in enumerate(sents):
            for k in ("a_exact", "b_variant"):
                for m in TIERS[k].finditer(s):
                    window = " ".join(sents[max(0, i - 1): i + 2])
                    ctx, review_flag = classify(window, s)
                    hits.append({"review_id": row.get("review_id"), "tier": k, "term": m.group(0),
                                 "context": ctx, "needs_review": review_flag,
                                 "rating": row.get("rating"), "sentence": s[:80],
                                 "polarity": polarity(row.get("rating"), window)})

    by_ctx = defaultdict(list)
    for h in hits:
        by_ctx[h["context"]].append(h)

    pct = lambda x: f"{x / n * 100:.1f}%" if n else "N/A (리뷰 0건)"
    out = ["## 단계별 집계", "", f"전체 리뷰 수: {n}", "",
           "| 단계 | 등장 횟수 | 해당 리뷰 수 | 전체 리뷰 대비 |", "|---|---|---|---|"]
    names = {"a_exact": "(a) '깊이' 정확 일치", "b_variant": "(b) 어근 변형", "c_opposite": "(c) 반대 표현"}
    for k, v in tier.items():
        out.append(f"| {names[k]} | {v['mentions']} | {v['reviews']} | {pct(v['reviews'])} |")
    out += ["", f"지정 목록 밖 '깊-' 활용형(미집계): {dict(other_forms) or '없음'}", "",
            "## 맥락별 결과 ((a)+(b) 등장 기준)", "",
            "| 맥락 | 건수 | 평점 평균 | 긍정 | 부정 | 검토 필요 | 대표 문장 |", "|---|---|---|---|---|---|---|"]
    for key in [r[0] for r in RULES] + ["other"]:
        hs = by_ctx.get(key, [])
        rs = [float(h["rating"]) for h in hs if str(h["rating"]).replace(".", "", 1).isdigit()]
        avg = f"{statistics.mean(rs):.2f}" if rs else "-"
        pos = sum(h["polarity"] == "pos" for h in hs); neg = sum(h["polarity"] == "neg" for h in hs)
        ratio = (lambda x: f"{x}/{len(hs)}") if hs else (lambda x: "-")
        rep = hs[0]["sentence"].replace("|", "/") if hs else "-"
        out.append(f"| {LABELS[key]} | {len(hs)} | {avg} | {ratio(pos)} | {ratio(neg)} | "
                   f"{sum(h['needs_review'] for h in hs)} | {rep} |")
    print("\n".join(out))
    if a.json:
        json.dump({"n_reviews": n, "tiers": tier, "other_forms": other_forms, "hits": hits},
                  open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

if __name__ == "__main__":
    sys.exit(main())
