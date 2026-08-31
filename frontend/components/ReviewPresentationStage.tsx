"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { useAppStore } from "@/lib/store";
import { ApiError, refineCopy, slideImageUrl, translatePresentation } from "@/lib/api";

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
  const currentPresIdx = useAppStore((s) => s.currentPresIdx);
  const setCurrentPresIdx = useAppStore((s) => s.setCurrentPresIdx);
  const setStage = useAppStore((s) => s.setStage);
  const resetReviewProgress = useAppStore((s) => s.resetReviewProgress);

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [chatMsgs, setChatMsgs] = useState<ChatMsg[]>([]);
  const [chatInput, setChatInput] = useState("");
  const [chatBusy, setChatBusy] = useState(false);
  const loadedRef = useRef(false);

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
      .then(setPresentationTranslations)
      .catch((e) => setError(e instanceof ApiError ? e.message : "번역 오류"))
      .finally(() => setLoading(false));
  }, [presentationUnits, keyPhrases, setPresentationTranslations]);

  useEffect(() => {
    setChatMsgs([]);
    setChatInput("");
  }, [slideIdx]);

  function goForward() {
    if (copyUnits.length) {
      resetReviewProgress();
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
        <div className="error-box" style={{ maxWidth: 1200, margin: "14px auto" }}>
          {error}
        </div>
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
          {slideUnits.map((unit) => {
            const item = presentationTranslations[unit.id] || { en_text: "", notes: "", clarification: "" };
            const note = item.notes || item.clarification;
            return (
              <div key={unit.id}>
                <div className="pair-block">
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
