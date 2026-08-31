"use client";

import { useAppStore } from "@/lib/store";
import type { Stage } from "@/lib/types";

const STAGES_KO_EN: Stage[] = ["upload", "classify", "review_2a", "review_2b", "download"];
const LABELS_KO_EN = ["파일 업로드", "슬라이드 분류", "발표용 감수", "카피 선택", "다운로드"];
const STAGES_EN_KO: Stage[] = ["upload", "en_ko", "download"];
const LABELS_EN_KO = ["파일 업로드", "번역", "다운로드"];

export default function Header() {
  const direction = useAppStore((s) => s.direction);
  const stage = useAppStore((s) => s.stage);

  const stages = direction === "en_ko" ? STAGES_EN_KO : STAGES_KO_EN;
  const labels = direction === "en_ko" ? LABELS_EN_KO : LABELS_KO_EN;
  const currIdx = stages.indexOf(stage);

  return (
    <>
      <div className="topbar">
        <div className="topbar-inner">
          <div className="topbar-logo">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <rect x="2" y="3" width="20" height="15" rx="2" stroke="white" strokeWidth="2" fill="none" />
              <path d="M8 22h8M12 18v4" stroke="white" strokeWidth="2" strokeLinecap="round" />
            </svg>
          </div>
          <span className="topbar-title">Agency Deck Translator</span>
        </div>
      </div>
      <div className="steprail">
        <div className="steprail-inner">
          {labels.map((label, i) => {
            const cls = i < currIdx ? "step done" : i === currIdx ? "step active" : "step";
            return (
              <div className={cls} key={label}>
                <span className="step-dot" />
                {label}
              </div>
            );
          })}
        </div>
      </div>
    </>
  );
}
