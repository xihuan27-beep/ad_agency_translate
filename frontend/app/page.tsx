"use client";

import { useEffect, useRef, useState } from "react";
import { useAppStore } from "@/lib/store";
import { createSession } from "@/lib/api";
import Header from "@/components/Header";
import UploadStage from "@/components/UploadStage";
import ClassifyStage from "@/components/ClassifyStage";
import ReviewPresentationStage from "@/components/ReviewPresentationStage";
import ReviewCopyStage from "@/components/ReviewCopyStage";
import EnKoStage from "@/components/EnKoStage";
import DownloadStage from "@/components/DownloadStage";

export default function Home() {
  const stage = useAppStore((s) => s.stage);
  const setSessionId = useAppStore((s) => s.setSessionId);
  const [error, setError] = useState("");
  const initRef = useRef(false);

  useEffect(() => {
    if (initRef.current) return;
    initRef.current = true;
    createSession()
      .then(setSessionId)
      .catch(() => setError("서버에 연결할 수 없습니다. 잠시 후 새로고침 해주세요."));
  }, [setSessionId]);

  return (
    <>
      <Header />
      {error && (
        <div className="page">
          <div className="error-box">{error}</div>
        </div>
      )}
      {!error && stage === "upload" && <UploadStage />}
      {!error && stage === "classify" && <ClassifyStage />}
      {!error && stage === "review_2a" && <ReviewPresentationStage />}
      {!error && stage === "review_2b" && <ReviewCopyStage />}
      {!error && stage === "en_ko" && <EnKoStage />}
      {!error && stage === "download" && <DownloadStage />}
    </>
  );
}
