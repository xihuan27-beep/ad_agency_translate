"use client";

import { useState } from "react";
import { reportError } from "@/lib/api";

export default function ErrorNotice({
  message,
  context,
  style,
}: {
  message: string;
  context?: string;
  style?: React.CSSProperties;
}) {
  const [status, setStatus] = useState<"idle" | "sending" | "sent" | "failed">("idle");

  async function handleReport() {
    setStatus("sending");
    try {
      await reportError(message, context || "");
      setStatus("sent");
    } catch {
      setStatus("failed");
    }
  }

  return (
    <div className="error-box" style={style}>
      <div>{message}</div>
      {status === "sent" ? (
        <div className="error-report-sent">신고가 접수되었습니다. 감사합니다.</div>
      ) : (
        <button className="error-report-btn" onClick={handleReport} disabled={status === "sending"}>
          {status === "sending" ? "전송 중..." : status === "failed" ? "전송 실패 - 다시 시도" : "오류 신고"}
        </button>
      )}
    </div>
  );
}
