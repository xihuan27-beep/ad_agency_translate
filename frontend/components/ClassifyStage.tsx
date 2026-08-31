"use client";

import { useMemo, useState } from "react";
import { useAppStore, materializeMerges } from "@/lib/store";
import { slideImageUrl } from "@/lib/api";
import type { Category, ClassifiedUnit } from "@/lib/types";

export default function ClassifyStage() {
  const sessionId = useAppStore((s) => s.sessionId);
  const hasSlideImages = useAppStore((s) => s.hasSlideImages);
  const slideCount = useAppStore((s) => s.slideCount);
  const classifiedUnits = useAppStore((s) => s.classifiedUnits);
  const toggleCategory = useAppStore((s) => s.toggleCategory);
  const setAllCategory = useAppStore((s) => s.setAllCategory);
  const excludedIds = useAppStore((s) => s.excludedIds);
  const toggleExcluded = useAppStore((s) => s.toggleExcluded);
  const mergedSeps = useAppStore((s) => s.mergedSeps);
  const toggleMerge = useAppStore((s) => s.toggleMerge);
  const addManualUnit = useAppStore((s) => s.addManualUnit);
  const activeSlide = useAppStore((s) => s.activeSlide);
  const setActiveSlide = useAppStore((s) => s.setActiveSlide);
  const setStage = useAppStore((s) => s.setStage);
  const setReviewUnits = useAppStore((s) => s.setReviewUnits);
  const resetReviewProgress = useAppStore((s) => s.resetReviewProgress);

  const [manualText, setManualText] = useState("");
  const [manualCat, setManualCat] = useState<Category>("presentation");
  const [error, setError] = useState("");

  const nSlides = useMemo(() => {
    const fromUnits = classifiedUnits.length
      ? Math.max(...classifiedUnits.map((u) => u.slide_idx)) + 1
      : 0;
    return Math.max(slideCount, fromUnits);
  }, [classifiedUnits, slideCount]);

  const slideGroups = useMemo(() => {
    const map = new Map<number, { i: number; u: ClassifiedUnit }[]>();
    classifiedUnits.forEach((u, i) => {
      const arr = map.get(u.slide_idx) || [];
      arr.push({ i, u });
      map.set(u.slide_idx, arr);
    });
    return map;
  }, [classifiedUnits]);

  const slideItems = slideGroups.get(activeSlide) || [];
  const mergedSet = mergedSeps[activeSlide] || new Set<number>();

  function handleNext() {
    const merged = materializeMerges(classifiedUnits, mergedSeps);
    const pres = merged.filter((u) => u.category === "presentation" && !excludedIds.has(u.id));
    const copy = merged.filter((u) => u.category === "copy" && !excludedIds.has(u.id));
    if (!pres.length && !copy.length) {
      setError("분류된 텍스트가 없습니다.");
      return;
    }
    setReviewUnits(pres, copy);
    resetReviewProgress();
    setStage(pres.length ? "review_2a" : "review_2b");
  }

  function handleManualAdd() {
    if (!manualText.trim()) return;
    addManualUnit({
      id: `manual_s${activeSlide}_${Date.now()}`,
      slide_idx: activeSlide,
      shape_id: -1,
      p_idx: 0,
      ko_text: manualText.trim(),
      font_size: 14,
      shape_text: manualText.trim(),
      shape_para_count: 1,
      category: manualCat,
    });
    setManualText("");
  }

  return (
    <>
      <div className="navrow">
        <button className="btn" onClick={() => setStage("upload")}>
          ← 업로드로 돌아가기
        </button>
        <div className="navrow-desc">Copywriter가 번역해야할 카피와 아닌 발표용을 구분할 수 있습니다</div>
        <button className="btn btn-primary" onClick={handleNext}>
          발표용 감수로 넘어가기 →
        </button>
      </div>

      {error && (
        <div className="error-box" style={{ maxWidth: 1200, margin: "0 auto 14px" }}>
          {error}
        </div>
      )}

      <div className="classify-3col">
        <div className="classify-thumbs">
          {Array.from({ length: nSlides }, (_, s) => (
            <div className="thumb-row" key={s}>
              <span className="thumb-num">{s + 1}</span>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div
                  className={`thumb-item${s === activeSlide ? " active" : ""}`}
                  onClick={() => setActiveSlide(s)}
                >
                  <div className="thumb-slide">
                    {hasSlideImages && sessionId ? (
                      <img src={slideImageUrl(sessionId, s)} alt={`슬라이드 ${s + 1}`} />
                    ) : null}
                  </div>
                </div>
                <div
                  className={`thumb-select${s === activeSlide ? " active" : ""}`}
                  onClick={() => setActiveSlide(s)}
                >
                  선택
                </div>
              </div>
            </div>
          ))}
        </div>

        <div className="classify-center">
          <div className="bezel">
            {hasSlideImages && sessionId ? (
              <img src={slideImageUrl(sessionId, activeSlide)} alt={`슬라이드 ${activeSlide + 1}`} />
            ) : (
              <div className="bezel-placeholder">슬라이드 {activeSlide + 1}</div>
            )}
          </div>
          <div className="bezel-caption">미리보기는 서버 폰트 제한으로 실제 PPT와 다를 수 있습니다</div>
        </div>

        <div className="classify-right">
          <div className="card">
            <div className="card-title">슬라이드 {activeSlide + 1}</div>
            {slideItems.length === 0 ? (
              <div style={{ fontSize: 13, color: "var(--cm)" }}>이 슬라이드에 번역할 텍스트가 없습니다.</div>
            ) : (
              <>
                <div className="bulk-row">
                  <button
                    className="bulk-btn pres"
                    onClick={() => setAllCategory(activeSlide, "presentation")}
                  >
                    전체 발표용
                  </button>
                  <button className="bulk-btn copy" onClick={() => setAllCategory(activeSlide, "copy")}>
                    전체 카피
                  </button>
                </div>

                {slideItems.map(({ u }, pos) => {
                  const isExcluded = excludedIds.has(u.id);
                  const short = u.ko_text.length > 45 ? u.ko_text.slice(0, 45) + "…" : u.ko_text;
                  const mergeAbove = mergedSet.has(pos - 1);
                  const mergeBelow = mergedSet.has(pos);
                  let cardCls = "item-card";
                  if (mergeAbove && mergeBelow) cardCls += " merge-mid";
                  else if (mergeAbove) cardCls += " merge-bottom";
                  else if (mergeBelow) cardCls += " merge-top";
                  if (isExcluded) cardCls += " excluded";
                  const hasNext = pos + 1 < slideItems.length;

                  return (
                    <div key={u.id}>
                      <div className={cardCls}>
                        {isExcluded ? (
                          <span style={{ fontSize: 11, color: "var(--ct3)", flexShrink: 0 }}>제외됨</span>
                        ) : (
                          <button
                            className={`item-tag ${u.category === "presentation" ? "pres" : "copy"}`}
                            onClick={() => toggleCategory(u.id)}
                          >
                            {u.category === "presentation" ? "발표용" : "카피"}
                          </button>
                        )}
                        <div className="item-text">{short}</div>
                        <button className="item-excl-btn" onClick={() => toggleExcluded(u.id)}>
                          {isExcluded ? "복원" : "✕"}
                        </button>
                      </div>
                      {hasNext && (
                        <div className="separator">
                          <div className={`sep-line${mergeBelow ? " faint" : ""}`} />
                          <button
                            className={`sep-btn${mergeBelow ? " merged" : ""}`}
                            title={mergeBelow ? "합쳐진 항목 나누기" : "다음 항목과 합치기"}
                            onClick={() => toggleMerge(activeSlide, pos)}
                          >
                            {mergeBelow ? "○" : "✕"}
                          </button>
                          <div className={`sep-line${mergeBelow ? " faint" : ""}`} />
                        </div>
                      )}
                    </div>
                  );
                })}
              </>
            )}
          </div>

          <div className="card">
            <div className="card-title">텍스트 직접 추가</div>
            <input
              className="field"
              style={{ marginBottom: 8 }}
              placeholder="인식되지 않은 텍스트를 입력하세요"
              value={manualText}
              onChange={(e) => setManualText(e.target.value)}
            />
            <div style={{ display: "flex", gap: 8 }}>
              <select
                className="field"
                style={{ flex: 2 }}
                value={manualCat === "presentation" ? "발표용" : "카피"}
                onChange={(e) => setManualCat(e.target.value === "발표용" ? "presentation" : "copy")}
              >
                <option value="발표용">발표용</option>
                <option value="카피">카피</option>
              </select>
              <button className="btn btn-primary" style={{ flex: 1 }} onClick={handleManualAdd}>
                추가
              </button>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}
