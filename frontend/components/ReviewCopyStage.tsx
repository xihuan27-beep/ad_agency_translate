"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { useAppStore } from "@/lib/store";
import { ApiError, grammarCheck, refineCopy, slideImageUrl, translateCopyOptions } from "@/lib/api";
import type { TextUnit } from "@/lib/types";

const OPTS_META: [string, string][] = [
  ["의역", "Feel, rhythm, impact 우선"],
  ["균형", "의미 + 자연스러운 영어"],
  ["직역", "원문 의미 중심"],
];

export default function ReviewCopyStage() {
  const sessionId = useAppStore((s) => s.sessionId);
  const hasSlideImages = useAppStore((s) => s.hasSlideImages);
  const copyUnits = useAppStore((s) => s.copyUnits);
  const presentationUnits = useAppStore((s) => s.presentationUnits);
  const keyPhrases = useAppStore((s) => s.keyPhrases);
  const copyOptions = useAppStore((s) => s.copyOptions);
  const setCopyOptions = useAppStore((s) => s.setCopyOptions);
  const copySelections = useAppStore((s) => s.copySelections);
  const setCopySelectionsForIds = useAppStore((s) => s.setCopySelectionsForIds);
  const setCopySelectionsMap = useAppStore((s) => s.setCopySelectionsMap);
  const updateCopyOptionText = useAppStore((s) => s.updateCopyOptionText);
  const currentCopyIdx = useAppStore((s) => s.currentCopyIdx);
  const setCurrentCopyIdx = useAppStore((s) => s.setCurrentCopyIdx);
  const setStage = useAppStore((s) => s.setStage);

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [grammarResults, setGrammarResults] = useState<Record<string, string>>({});
  const [grammarBusy, setGrammarBusy] = useState(false);
  const [chatInput, setChatInput] = useState("");
  const [chatBusy, setChatBusy] = useState(false);
  const loadedRef = useRef(false);

  useEffect(() => {
    if (loadedRef.current || !copyUnits.length) return;
    loadedRef.current = true;
    setLoading(true);
    translateCopyOptions(copyUnits, keyPhrases)
      .then((opts) => {
        setCopyOptions(opts);
        const initial: Record<string, string> = {};
        for (const u of copyUnits) {
          initial[u.id] = opts[u.id]?.options?.[0] || "";
        }
        setCopySelectionsMap(initial);
      })
      .catch((e) => setError(e instanceof ApiError ? e.message : "카피 옵션 생성 오류"))
      .finally(() => setLoading(false));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [copyUnits, keyPhrases]);

  const groups = useMemo(() => {
    const map = new Map<string, TextUnit[]>();
    for (const u of copyUnits) {
      const arr = map.get(u.ko_text) || [];
      arr.push(u);
      map.set(u.ko_text, arr);
    }
    return Array.from(map.values());
  }, [copyUnits]);

  const total = groups.length;
  const idx = Math.min(currentCopyIdx, Math.max(total - 1, 0));
  const group = groups[idx] || [];
  const unit = group[0];

  useEffect(() => {
    setChatInput("");
  }, [idx]);

  if (loading) {
    return (
      <div className="spinner-wrap">
        <div className="spinner" />
        AI가 카피 옵션을 생성하는 중... ({copyUnits.length}개 카피 항목)
      </div>
    );
  }

  if (!total || !unit) {
    return (
      <div className="page page-narrow">
        {error && <div className="error-box">{error}</div>}
        <div className="card">카피 텍스트가 없습니다.</div>
        <button className="btn btn-primary btn-block" onClick={() => setStage("download")}>
          다운로드 →
        </button>
      </div>
    );
  }

  const data = copyOptions[unit.id] || { options: ["", "", ""] as [string, string, string], notes: "", recommendation: "", cultural_flag: "", clarification: "" };
  const options = data.options;
  const currentSel = copySelections[unit.id] ?? options[0] ?? "";
  const dupSlides = Array.from(new Set(group.map((u) => u.slide_idx))).sort((a, b) => a - b);
  const groupIds = group.map((u) => u.id);
  const noteContent = data.notes || data.clarification || data.cultural_flag;
  const grammarFeedback = grammarResults[unit.id] || "";

  async function handleUseOption(optText: string) {
    setCopySelectionsForIds(groupIds, optText);
  }

  async function handleGrammarCheck() {
    setGrammarBusy(true);
    try {
      const feedback = await grammarCheck(unit.ko_text, currentSel, keyPhrases);
      setGrammarResults((r) => ({ ...r, [unit.id]: feedback }));
    } catch (e) {
      setGrammarResults((r) => ({ ...r, [unit.id]: e instanceof ApiError ? e.message : "오류가 발생했습니다." }));
    } finally {
      setGrammarBusy(false);
    }
  }

  async function handleChatSend() {
    const msg = chatInput.trim();
    if (!msg || chatBusy) return;
    setChatBusy(true);
    try {
      const refined = await refineCopy(unit.ko_text, currentSel, msg, keyPhrases);
      setCopySelectionsForIds(groupIds, refined);
      setChatInput("");
    } catch (e) {
      setError(e instanceof ApiError ? e.message : "오류가 발생했습니다.");
    } finally {
      setChatBusy(false);
    }
  }

  return (
    <>
      <div className="review-layout">
        <div className="review-left">
          <div className="bezel">
            {hasSlideImages && sessionId ? (
              <img src={slideImageUrl(sessionId, unit.slide_idx)} alt={`슬라이드 ${unit.slide_idx + 1}`} />
            ) : (
              <div className="bezel-placeholder">슬라이드 {unit.slide_idx + 1}</div>
            )}
          </div>
          <div className="bezel-caption">
            슬라이드 {unit.slide_idx + 1}
            {dupSlides.length > 1 ? ` (슬라이드 ${dupSlides.map((s) => s + 1).join(", ")}에 반복)` : ""}
          </div>
          <div className="bezel-caption">미리보기는 서버 폰트 제한으로 실제 PPT와 다를 수 있습니다</div>
        </div>

        <div className="review-right">
          {error && <div className="error-box">{error}</div>}

          <div className="pair-block">
            <div className="ko-block">{unit.ko_text}</div>
          </div>

          {dupSlides.length > 1 && (
            <div className="notebox">
              🔁 슬라이드 {dupSlides.map((s) => s + 1).join(", ")}에 동일 카피 — 한 번 선택하면 모두 적용됩니다
            </div>
          )}

          {noteContent && <div className="notebox">📝 {noteContent}</div>}

          {OPTS_META.map(([name, sub], i) => {
            const optText = options[i] || "";
            const isSel = currentSel === optText;
            return (
              <div className={`copyopt${isSel ? " selected" : ""}`} key={name}>
                <div className="opt-label-row">
                  <span className="opt-tag">{name}</span>
                  <span className="opt-sub">{sub}</span>
                </div>
                <textarea
                  className="opt-text"
                  rows={2}
                  value={optText}
                  onChange={(e) => updateCopyOptionText(unit.id, i, e.target.value)}
                />
                {isSel ? (
                  <div className="opt-selected-label">✓ 선택됨</div>
                ) : (
                  <button className="btn opt-use-btn btn-block" onClick={() => handleUseOption(optText)}>
                    이 버전 사용
                  </button>
                )}
              </div>
            );
          })}

          {data.recommendation && (
            <div className="recbox">
              📌 <strong>추천 이유:</strong> {data.recommendation}
            </div>
          )}

          <button className="btn" onClick={handleGrammarCheck} disabled={grammarBusy}>
            {grammarBusy ? "체크 중..." : "선택된 카피 문법 체크"}
          </button>
          {grammarFeedback && (
            <div className="notebox" style={{ background: "#F0FDF4", border: "1px solid #86EFAC", color: "#166534" }}>
              ✅ {grammarFeedback}
            </div>
          )}

          <div className="chat-section">
            <div className="chat-header">AI 수정 요청</div>
            <div className="chat-input-row">
              <input
                className="field"
                placeholder={`수정 요청 내용을 입력하세요 (${idx + 1}/${total})`}
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
      </div>

      <div className="stepnav">
        <button className="btn" disabled={idx === 0} onClick={() => setCurrentCopyIdx(idx - 1)}>
          ＜
        </button>
        <div className="stepnav-count">
          {idx + 1}/{total} 카피
        </div>
        <button className="btn" disabled={idx >= total - 1} onClick={() => setCurrentCopyIdx(idx + 1)}>
          ＞
        </button>
      </div>

      <div className="stagenav">
        <button className="btn" onClick={() => setStage(presentationUnits.length ? "review_2a" : "classify")}>
          ← 발표용 감수로
        </button>
        <button className="btn btn-primary" onClick={() => setStage("download")}>
          다음 단계: 다운로드 →
        </button>
      </div>
    </>
  );
}
