"use client";

import { useEffect, useRef, useState } from "react";
import { useAppStore } from "@/lib/store";
import { ApiError, applyTranslations, downloadUrl } from "@/lib/api";

export default function DownloadStage() {
  const sessionId = useAppStore((s) => s.sessionId);
  const direction = useAppStore((s) => s.direction);
  const fileType = useAppStore((s) => s.fileType);
  const textUnits = useAppStore((s) => s.textUnits);
  const enKoTranslations = useAppStore((s) => s.enKoTranslations);
  const presentationUnits = useAppStore((s) => s.presentationUnits);
  const copyUnits = useAppStore((s) => s.copyUnits);
  const presentationTranslations = useAppStore((s) => s.presentationTranslations);
  const copySelections = useAppStore((s) => s.copySelections);
  const reset = useAppStore((s) => s.reset);

  const [ready, setReady] = useState(false);
  const [error, setError] = useState("");
  const loadedRef = useRef(false);

  const isEnKo = direction === "en_ko";

  useEffect(() => {
    if (loadedRef.current || !sessionId) return;
    loadedRef.current = true;
    let translations: Record<string, string>;
    if (isEnKo) {
      translations = enKoTranslations;
    } else {
      const presTrans: Record<string, string> = {};
      for (const [uid, v] of Object.entries(presentationTranslations)) presTrans[uid] = v.en_text;
      translations = { ...presTrans, ...copySelections };
    }
    applyTranslations(sessionId, translations, "")
      .then(() => setReady(true))
      .catch((e) => setError(e instanceof ApiError ? e.message : "파일 생성 오류"));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sessionId]);

  const subtitle = isEnKo
    ? `영어 원문 ${textUnits.length}개 항목 → 한국어 번역 ${Object.values(enKoTranslations).filter((v) => v.trim()).length}개 완료`
    : `발표용 텍스트 ${presentationUnits.length}개 번역 · 광고 카피 ${copyUnits.length}개 트랜스크리에이션 완료`;

  const outputFormatLabel = fileType === "pptx" ? "PPT" : "Word";
  const dlLabel = isEnKo ? `한국어 ${outputFormatLabel} 다운로드` : `영문 ${outputFormatLabel} 다운로드`;

  function handleReset() {
    if (sessionId) reset();
    window.location.reload();
  }

  return (
    <div className="download-wrap">
      <div className="dl-icon">✅</div>
      <div className="dl-title">번역 완료</div>
      <div className="dl-sub">{subtitle}</div>

      {error && <div className="error-box">{error}</div>}

      <div className="dl-actions">
        {ready && sessionId ? (
          <a className="btn btn-primary btn-block" href={downloadUrl(sessionId)}>
            {dlLabel}
          </a>
        ) : !error ? (
          <div className="spinner-wrap" style={{ padding: "20px 0" }}>
            <div className="spinner" />
            번역본을 파일에 적용하는 중...
          </div>
        ) : null}
        <button className="btn" onClick={handleReset}>
          처음으로 돌아가기
        </button>
      </div>
    </div>
  );
}
