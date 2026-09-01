"use client";

import { useState } from "react";
import { useAppStore } from "@/lib/store";
import { ApiError, classify, fetchFile, translateEnToKo } from "@/lib/api";

export default function UploadStage() {
  const sessionId = useAppStore((s) => s.sessionId);
  const direction = useAppStore((s) => s.direction);
  const setDirection = useAppStore((s) => s.setDirection);
  const fetchedUrl = useAppStore((s) => s.fetchedUrl);
  const setFetchedUrl = useAppStore((s) => s.setFetchedUrl);
  const setUploadResult = useAppStore((s) => s.setUploadResult);
  const keyPhrases = useAppStore((s) => s.keyPhrases);
  const setKeyPhrases = useAppStore((s) => s.setKeyPhrases);
  const textUnits = useAppStore((s) => s.textUnits);
  const setClassifiedUnits = useAppStore((s) => s.setClassifiedUnits);
  const setEnKoTranslations = useAppStore((s) => s.setEnKoTranslations);
  const setStage = useAppStore((s) => s.setStage);

  const [urlInput, setUrlInput] = useState("");
  const [fetching, setFetching] = useState(false);
  const [starting, setStarting] = useState(false);
  const [error, setError] = useState("");

  const isEnKo = direction === "en_ko";
  const urlTrimmed = urlInput.trim();
  const alreadyFetched = !!urlTrimmed && urlTrimmed === fetchedUrl;
  const fileReady = alreadyFetched && textUnits.length > 0;

  async function handleFetch() {
    if (!sessionId || !urlTrimmed || alreadyFetched) return;
    setError("");
    setFetching(true);
    try {
      const result = await fetchFile(sessionId, urlTrimmed);
      setUploadResult(result);
      setFetchedUrl(urlTrimmed);
    } catch (e) {
      setError(e instanceof ApiError ? e.message : "다운로드 실패");
    } finally {
      setFetching(false);
    }
  }

  async function handleStart() {
    if (!sessionId || !fileReady) return;
    setError("");
    setStarting(true);
    try {
      if (isEnKo) {
        const translations = await translateEnToKo(textUnits, keyPhrases);
        setEnKoTranslations(translations);
        setStage("en_ko");
      } else {
        const classified = await classify(sessionId, keyPhrases);
        setClassifiedUnits(classified);
        setStage("classify");
      }
    } catch (e) {
      setError(e instanceof ApiError ? e.message : "오류가 발생했습니다");
    } finally {
      setStarting(false);
    }
  }

  function updateTerm(idx: number, lang: "한국어" | "영어", value: string) {
    const next = keyPhrases.map((kp, i) => (i === idx ? { ...kp, [lang]: value } : kp));
    setKeyPhrases(next);
  }

  const colOrder: Array<["한국어" | "영어", string]> = isEnKo
    ? [
        ["영어", "English"],
        ["한국어", "한국어"],
      ]
    : [
        ["한국어", "한국어"],
        ["영어", "English"],
      ];

  return (
    <div className="page page-narrow">
      <div className="section-eyebrow">번역 선택</div>
      <div className="dir-grid" style={{ marginBottom: 20 }}>
        <button
          className={direction === "ko_en" ? "btn btn-primary btn-block" : "btn btn-block"}
          onClick={() => setDirection("ko_en")}
        >
          한국어 → 영어
        </button>
        <button
          className={direction === "en_ko" ? "btn btn-primary btn-block" : "btn btn-block"}
          onClick={() => setDirection("en_ko")}
        >
          영어 → 한국어
        </button>
      </div>

      {error && <div className="error-box">{error}</div>}

      <div className="card">
        <div className="card-title">번역할 PPT 파일 업로드 하기</div>
        <div style={{ display: "flex", gap: 8 }}>
          <input
            className="field"
            style={{ flex: 5 }}
            placeholder="https://drive.google.com/file/d/..."
            value={urlInput}
            onChange={(e) => setUrlInput(e.target.value)}
          />
          <button
            className="btn btn-primary"
            style={{ flex: 1 }}
            disabled={alreadyFetched || !urlTrimmed || fetching}
            onClick={handleFetch}
          >
            {fetching ? "..." : "가져오기"}
          </button>
        </div>
        <div style={{ fontSize: 11.5, color: "var(--cm)", marginTop: 8, lineHeight: 1.7 }}>
          무료버전에서는 Google Drive만 가능합니다 (PPT, Word, PDF 지원)
          <br />
          링크 공유 시 권한을 편집자(edit)로 해야 합니다
        </div>
      </div>

      <div className="card">
        <div className="card-title">번역 퀄리티 상승을 위한 추가 정보 입력</div>
        <div className="card-sub">
          브랜드명, 제품명, 내부 약어, 선호 문구 등 주요 용어를 등록하시면 이를 활용하여 번역하여 검수 업무가
          줄어듭니다
        </div>
        {keyPhrases.map((kp, idx) => (
          <div className="term-grid" key={idx}>
            {colOrder.map(([lang, placeholder]) => (
              <input
                key={lang}
                className="field"
                placeholder={placeholder}
                value={kp[lang]}
                onChange={(e) => updateTerm(idx, lang, e.target.value)}
              />
            ))}
          </div>
        ))}
        <button className="btn" onClick={() => setKeyPhrases([...keyPhrases, { 한국어: "", 영어: "" }])}>
          + 항목 추가
        </button>
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end" }}>
        <button
          className="btn btn-primary"
          style={{ width: "50%" }}
          disabled={!fileReady || starting}
          onClick={handleStart}
        >
          {starting ? "처리 중..." : "시작하기"}
        </button>
      </div>
    </div>
  );
}
