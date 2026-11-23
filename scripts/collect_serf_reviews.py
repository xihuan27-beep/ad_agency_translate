"""unni.app '세르프' 검색 결과에서 로그인 없이 공개된 리뷰만 수집한다.

원칙
- robots.txt를 먼저 읽고, 허용되지 않은 경로는 요청하지 않는다.
- 요청 간격 1초 이상. HTTP 429/403이 나오면 즉시 중단한다.
- 로그인/블러/'앱에서 보기' 가드가 켜져 있으면 우회하지 않고 그대로 보고한다.
- 닉네임·유저ID·사진은 저장하지 않는다. 남기는 필드: 본문, 평점, 시술명, 병원명, 날짜.

사용법: python scripts/collect_serf_reviews.py [--out reviews/serf_reviews.csv]
"""
import argparse, csv, json, re, sys, time, urllib.parse, urllib.request, urllib.robotparser

BASE = "https://unni.app"
QUERY = "세르프"
# 사용자가 지시한 1회성 조회이므로 robots.txt의 Claude-User 규칙을 적용한다.
# 이 그룹에 없는 범용 크롤러는 'User-agent: *  Disallow: /' 로 전면 차단된다.
UA = "Claude-User"
MIN_INTERVAL = 1.5
FIELDS = ["review_id", "date", "rating", "treatment", "hospital", "body"]

_last = 0.0

class Stop(Exception):
    pass

def fetch(url, rp):
    global _last
    if not rp.can_fetch(UA, url):
        raise Stop(f"robots.txt가 {UA}에게 허용하지 않는 경로: {url}")
    wait = MIN_INTERVAL - (time.time() - _last)
    if wait > 0:
        time.sleep(wait)
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            _last = time.time()
            return r.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        _last = time.time()
        if e.code in (403, 429):
            raise Stop(f"HTTP {e.code}, 수집 중단: {url}")
        raise

def next_data(html):
    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', html, re.S)
    return json.loads(m.group(1))["props"]["pageProps"] if m else {}

def pick(d, *keys):
    for k in keys:
        cur = d
        for part in k.split("."):
            cur = cur.get(part) if isinstance(cur, dict) else None
        if cur not in (None, ""):
            return cur
    return ""

def to_row(rv):
    """개인정보 필드(작성자, 닉네임, 사진, userId)는 일부러 읽지 않는다."""
    treat = pick(rv, "treatmentName", "treatment.name")
    if not treat and isinstance(rv.get("treatments"), list):
        treat = ";".join(t.get("name", "") for t in rv["treatments"] if isinstance(t, dict))
    return {
        "review_id": pick(rv, "id"),
        "date": pick(rv, "createTime", "createdAt", "date"),
        "rating": pick(rv, "rating", "score", "rating.amount"),
        "treatment": treat,
        "hospital": pick(rv, "hospital.name", "hospitalName"),
        "body": pick(rv, "contents", "content", "body", "text"),
    }

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="reviews/serf_reviews.csv")
    ap.add_argument("--report", default="reviews/serf_collection_log.json")
    ap.add_argument("--compare", nargs="*", default=[],
                    help="비교용 다른 시술 검색어. 공개 리뷰 수/가드 여부만 기록(본문 저장 안 함)")
    a = ap.parse_args()

    rp = urllib.robotparser.RobotFileParser(BASE + "/robots.txt")
    rp.read()
    log = {"user_agent": UA, "robots": {
        p: rp.can_fetch(UA, BASE + p) for p in ["/search", "/reviews", "/reviews/1", "/community", "/hospitals/1"]
    }, "robots_wildcard_all_blocked": not rp.can_fetch("*", BASE + "/search")}

    rows, stop_reason = [], None
    url = f"{BASE}/search?q={urllib.parse.quote(QUERY)}"
    try:
        pp = next_data(fetch(url, rp))
        log["review_total_reported"] = pp.get("reviewTotalLength")
        log["review_guard_active"] = pp.get("isReviewGuardActive")
        public = pp.get("reviews") or []
        log["reviews_public_on_page"] = len(public)
        rows = [to_row(r) for r in public if isinstance(r, dict)]
        if pp.get("isReviewGuardActive"):
            stop_reason = "리뷰 가드 활성화('로그인하고 보기'). 우회하지 않음."
        # 리뷰 상세/목록(/reviews)은 robots.txt가 막고 있어 페이지네이션을 시도하지 않는다.
        if not rp.can_fetch(UA, BASE + "/reviews"):
            stop_reason = (stop_reason or "") + " /reviews 경로는 robots.txt 차단."
    except Stop as e:
        stop_reason = str(e)

    log["compare"] = {}
    for q in a.compare:
        if stop_reason and "HTTP" in stop_reason:
            break
        try:
            cp = next_data(fetch(f"{BASE}/search?q={urllib.parse.quote(q)}", rp))
            log["compare"][q] = {"review_total_reported": cp.get("reviewTotalLength"),
                                 "reviews_public_on_page": len(cp.get("reviews") or []),
                                 "review_guard_active": cp.get("isReviewGuardActive")}
        except Stop as e:
            log["compare"][q] = {"stopped": str(e)}
            break

    with open(a.out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(rows)
    total = log.get("review_total_reported") or 0
    log.update(collected=len(rows), collection_rate=(len(rows) / total if total else None),
               stop_reason=stop_reason, collected_at=time.strftime("%Y-%m-%dT%H:%M:%S%z"))
    with open(a.report, "w", encoding="utf-8") as f:
        json.dump(log, f, ensure_ascii=False, indent=2)
    print(json.dumps(log, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    sys.exit(main())
