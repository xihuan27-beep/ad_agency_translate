"use client";

import { useAppStore } from "@/lib/store";

export default function EnKoStage() {
  const textUnits = useAppStore((s) => s.textUnits);
  const enKoTranslations = useAppStore((s) => s.enKoTranslations);
  const updateEnKoTranslation = useAppStore((s) => s.updateEnKoTranslation);
  const setStage = useAppStore((s) => s.setStage);

  const rows = textUnits.filter((u) => u.ko_text.trim());
  const nDone = rows.filter((u) => (enKoTranslations[u.id] || "").trim()).length;

  return (
    <div className="page">
      <div className="section-eyebrow">번역 검토</div>
      <div className="card-sub" style={{ marginBottom: 14 }}>
        총 {rows.length}개 텍스트 중 {nDone}개 번역 완료
      </div>

      <div className="card" style={{ maxHeight: 560, overflowY: "auto" }}>
        {rows.map((u) => (
          <div className="pair-block" key={u.id} style={{ marginBottom: 10 }}>
            <div className="ko-block">{u.ko_text}</div>
            <div className="en-block">
              <textarea
                className="en-text"
                rows={2}
                value={enKoTranslations[u.id] || ""}
                onChange={(e) => updateEnKoTranslation(u.id, e.target.value)}
              />
            </div>
          </div>
        ))}
      </div>

      <div className="stagenav">
        <button className="btn" onClick={() => setStage("upload")}>
          ← 업로드로 돌아가기
        </button>
        <button className="btn btn-primary" onClick={() => setStage("download")}>
          한국어 파일 생성 및 다운로드 →
        </button>
      </div>
    </div>
  );
}
