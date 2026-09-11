"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { useAppStore } from "@/lib/store";
import { ApiError, refineCopy, reviewPresentation, slideImageUrl, translatePresentation } from "@/lib/api";
import ErrorNotice from "@/components/ErrorNotice";

interface ChatMsg {
  role: "user" | "ai";
  text: string;
}

export default function ReviewPresentationStage() {
  const sessionId = useAppStore((s) => s.sessionId);
  const hasSlideImages = useAppStore((s) => s.hasSlideImages);
  const presentationUnits = useAppStore((s) => s.presentationUnits);
  const copyUnits = useAppStore((s) => s.copyUnits);
  const keyPhrases = useAppStore((s) => s.keyPhrases);
  const presentationTranslations = useAppStore((s) => s.presentationTranslations);
  const setPresentationTranslations = useAppStore((s) => s.setPresentationTranslations);
  const updatePresentationTranslation = useAppStore((s) => s.updatePresentationTranslation);
  const presentationReview = useAppStore((s) => s.presentationReview);
  const setPresentationReview = useAppStore((s) => s.setPresentationReview);
  const currentPresIdx = useAppStore((s) => s.currentPresIdx);
  const setCurrentPresIdx = useAppStore((s) => s.setCurrentPresIdx);
  const setStage = useAppStore((s) => s.setStage);

  const [loading, setLoading] = useState(false);
  const [reviewLoading, setReviewLoading] = useState(false);
  const [error, setError] = useState("");
  const [chatMsgs, setChatMsgs] = useState<ChatMsg[]>([]);
  const [chatInput, setChatInput] = useState("");
  const [chatBusy, setChatBusy] = useState(false);
  const loadedRef = useRef(false);
  const reviewLoadedRef = useRef(false);

  const slideBySlide = useMemo(() => {
    const map = new Map<number, typeof presentationUnits>();
    for (const u of presentationUnits) {
      const arr = map.get(u.slide_idx) || [];
      arr.push(u);
      map.set(u.slide_idx, arr);
    }
    return map;
  }, [presentationUnits]);
  const slideKeys = useMemo(() => Array.from(slideBySlide.keys()).sort((a, b) => a - b), [slideBySlide]);
  const total = slideKeys.length;
  const slidePos = Math.min(currentPresIdx, Math.max(total - 1, 0));
  const slideIdx = slideKeys[slidePos];
  const slideUnits = slideBySlide.get(slideIdx) || [];

  useEffect(() => {
    if (loadedRef.current || !presentationUnits.length) return;
    loadedRef.current = true;
    setLoading(true);
    translatePresentation(presentationUnits, keyPhrases)
      .then((translations) => {
        setPresentationTranslations(translations);
        if (reviewLoadedRef.current) return;
        reviewLoadedRef.current = true;
        const enTexts: Record<string, string> = {};
        for (const [id, v] of Object.entries(translations)) enTexts[id] = v.en_text;
        setReviewLoading(true);
        reviewPresentation(presentationUnits, enTexts, keyPhrases)
          .then(setPresentationReview)
          .catch(() => {})
          .finally(() => setReviewLoading(false));
      })
      .catch((e) => setError(e instanceof ApiError ? e.message : "번역 오류"))
      .finally(() => setLoading(false));
  }, [presentationUnits, keyPhrases, setPresentationTranslations, setPresentationReview]);

  const flaggedSlidePositions = useMemo(() => {
    const positions: number[] = [];
    slideKeys.forEach((sIdx, pos) => {
      const units = slideBySlide.get(sIdx) || [];
      if (units.some((u) => presentationReview[u.id]?.flagged)) positions.push(pos);
    });
    return positions;
  }, [slideKeys, slideBySlide, presentationReview]);

  const flaggedCount = useMemo(
    () => presentationUnits.filter((u) => presentationReview[u.id]?.flagged).length,
    [presentationUnits, presentationReview]
  );

  function goToNextFlagged() {
    if (!flaggedSlidePositions.length) return;
    const next = flaggedSlidePositions.find((p) => p > slidePos) ?? flaggedSlidePositions[0];
    setCurrentPresIdx(next);
  }

  useEffect(() => {
    setChatMsgs([]);
    setChatInput("");
  }, [slideIdx]);

  function goForward() {
    if (copyUnits.length) {
      setStage("review_2b");
    } else {
      setStage("download");
    }
  }

  async function handleChatSend() {
    const msg = chatInput.trim();
    if (!msg || chatBusy) return;
    setChatMsgs((m) => [...m, { role: "user", text: msg }]);
    setChatInput("");
    setChatBusy(true);
    try {
      for (const unit of slideUnits) {
        const item = presentationTranslations[unit.id];
        const enText = item?.en_text || "";
        const refined = await refineCopy(unit.ko_text, enText, msg, keyPhrases);
        updatePresentationTranslation(unit.id, refined);
      }
      setChatMsgs((m) => [...m, { role: "ai", text: "수정을 반영했습니다." }]);
    } catch (e) {
      setChatMsgs((m) => [...m, { role: "ai", text: e instanceof ApiError ? e.message : "오류가 발생했습니다." }]);
    } finally {
      setChatBusy(false);
    }
  }

  if (loading) {
    return (
      <div className="spinner-wrap">
        <div className="spinner" />
        AI가 발표용 텍스트를 일괄 번역하는 중...
      </div>
    );
  }

  if (!presentationUnits.length) {
    return (
      <div className="page page-narrow">
        <div className="card">발표용 텍스트가 없습니다.</div>
        <button className="btn btn-primary btn-block" onClick={goForward}>
          다음 →
        </button>
      </div>
    );
  }

  return (
    <>
      {error && (
        <ErrorNotice message={error} context="발표용 감수 단계" style={{ maxWidth: 1200, margin: "14px auto" }} />
      )}
      <div className="review-layout">
        <div className="review-left">
          <div className="bezel">
            {hasSlideImages && sessionId ? (
              <img src={slideImageUrl(sessionId, slideIdx)} alt={`슬라이드 ${slideIdx + 1}`} />
            ) : (
              <div className="bezel-placeholder">슬라이드 {slideIdx + 1}</div>
            )}
          </div>
          <div className="bezel-caption">미리보기는 서버 폰트 제한으로 실제 PPT와 다를 수 있습니다</div>

          <div className="chat-section" style={{ marginTop: 16 }}>
            <div className="chat-header">AI 수정 요청</div>
            <div className="chat-messages">
              {chatMsgs.length === 0 && <div className="chat-empty">이 슬라이드 번역에 대한 수정을 요청해보세요</div>}
              {chatMsgs.map((m, i) => (
                <div className={`chat-msg ${m.role}`} key={i}>
                  {m.text}
                </div>
              ))}
            </div>
            <div className="chat-input-row">
              <input
                className="field"
                placeholder={`이 슬라이드 번역 수정 요청 (${slidePos + 1}/${total} 슬라이드)`}
                value={chatInput}
                onChange={(e) => setChatInput(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleChatSend()}
                disabled={chatBusy}
              />
              <button className="btn btn-primary" onClick={handleChatSend} disabled={chatBusy}>
                전송
              </button>
            </div>
          </div>
        </div>

        <div className="review-right">
          {(reviewLoading || flaggedCount > 0) && (
            <div className="review-flag-summary">
              {reviewLoading
                ? "AI가 번역 퀄리티를 검토하는 중..."
                : `⚠ AI가 검토를 권장하는 항목 ${flaggedCount}개 / 전체 ${presentationUnits.length}개`}
              {!reviewLoading && flaggedSlidePositions.length > 0 && (
                <button className="review-flag-next-btn" onClick={goToNextFlagged}>
                  다음 검토 필요 슬라이드로 →
                </button>
              )}
            </div>
          )}
          {slideUnits.map((unit) => {
            const item = presentationTranslations[unit.id] || { en_text: "", notes: "", clarification: "" };
            const note = item.notes || item.clarification;
            const review = presentationReview[unit.id];
            return (
              <div key={unit.id}>
                <div className={`pair-block${review?.flagged ? " flagged" : ""}`}>
                  <div className="ko-block">{unit.ko_text}</div>
                  <div className="en-block">
                    <textarea
                      className="en-text"
                      rows={3}
                      value={item.en_text}
                      onChange={(e) => updatePresentationTranslation(unit.id, e.target.value)}
                    />
                  </div>
                </div>
                {review?.flagged && <div className="pair-note flagged">⚠ {review.issue}</div>}
                {note && <div className="pair-note">📝 {note}</div>}
              </div>
            );
          })}
        </div>
      </div>

      <div className="stepnav">
        <button className="btn" disabled={slidePos === 0} onClick={() => setCurrentPresIdx(slidePos - 1)}>
          ＜
        </button>
        <div className="stepnav-count">
          {slidePos + 1}/{total} 슬라이드
        </div>
        <button className="btn" disabled={slidePos >= total - 1} onClick={() => setCurrentPresIdx(slidePos + 1)}>
          ＞
        </button>
      </div>

      <div className="stagenav">
        <button className="btn" onClick={() => setStage("classify")}>
          ← 분류로 돌아가기
        </button>
        <button className="btn btn-primary" onClick={goForward}>
          {copyUnits.length ? `다음 단계: 카피 선택 (${copyUnits.length}개) →` : "다음 단계: 다운로드 →"}
        </button>
      </div>
    </>
  );
}
