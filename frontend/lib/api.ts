import type { ClassifiedUnit, CopyOptionSet, KeyPhrasePair, PresentationTranslation, TextUnit } from "./types";

const API_BASE = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";

class ApiError extends Error {}

async function request<T>(path: string, options: RequestInit = {}): Promise<T> {
  const res = await fetch(`${API_BASE}${path}`, {
    ...options,
    headers: { "Content-Type": "application/json", ...(options.headers || {}) },
  });
  if (!res.ok) {
    let detail = res.statusText;
    try {
      const body = await res.json();
      detail = body.detail || detail;
    } catch {
      // ignore
    }
    throw new ApiError(detail);
  }
  return res.json();
}

export function slideImageUrl(sessionId: string, idx: number): string {
  return `${API_BASE}/api/sessions/${sessionId}/slide-image/${idx}.png`;
}

export function downloadUrl(sessionId: string): string {
  return `${API_BASE}/api/sessions/${sessionId}/download`;
}

export async function createSession(): Promise<string> {
  const data = await request<{ sessionId: string }>("/api/sessions", { method: "POST" });
  return data.sessionId;
}

export interface FetchResult {
  fileType: "pptx" | "docx";
  slideCount: number;
  hasSlideImages: boolean;
  textUnits: TextUnit[];
}

export async function fetchFile(sessionId: string, url: string): Promise<FetchResult> {
  return request(`/api/sessions/${sessionId}/fetch`, {
    method: "POST",
    body: JSON.stringify({ url }),
  });
}

export async function classify(
  sessionId: string,
  keyPhrases: KeyPhrasePair[]
): Promise<ClassifiedUnit[]> {
  const data = await request<{ units: ClassifiedUnit[] }>(`/api/sessions/${sessionId}/classify`, {
    method: "POST",
    body: JSON.stringify({ keyPhrases }),
  });
  return data.units;
}

export async function translatePresentation(
  units: TextUnit[],
  keyPhrases: KeyPhrasePair[]
): Promise<Record<string, PresentationTranslation>> {
  const data = await request<{ translations: Record<string, PresentationTranslation> }>(
    "/api/translate/presentation",
    { method: "POST", body: JSON.stringify({ units, keyPhrases }) }
  );
  return data.translations;
}

export async function translateCopyOptions(
  units: TextUnit[],
  keyPhrases: KeyPhrasePair[]
): Promise<Record<string, CopyOptionSet>> {
  const data = await request<{ options: Record<string, CopyOptionSet> }>(
    "/api/translate/copy-options",
    { method: "POST", body: JSON.stringify({ units, keyPhrases }) }
  );
  return data.options;
}

export async function refineCopy(
  koText: string,
  currentEn: string,
  instruction: string,
  keyPhrases: KeyPhrasePair[]
): Promise<string> {
  const data = await request<{ text: string }>("/api/translate/refine-copy", {
    method: "POST",
    body: JSON.stringify({ koText, currentEn, instruction, keyPhrases }),
  });
  return data.text;
}

export async function grammarCheck(
  koText: string,
  enText: string,
  keyPhrases: KeyPhrasePair[]
): Promise<string> {
  const data = await request<{ feedback: string }>("/api/translate/grammar-check", {
    method: "POST",
    body: JSON.stringify({ koText, enText, keyPhrases }),
  });
  return data.feedback;
}

export async function translateEnToKo(
  units: TextUnit[],
  keyPhrases: KeyPhrasePair[]
): Promise<Record<string, string>> {
  const data = await request<{ translations: Record<string, string> }>(
    "/api/translate/en-to-ko",
    { method: "POST", body: JSON.stringify({ units, keyPhrases }) }
  );
  return data.translations;
}

export async function applyTranslations(
  sessionId: string,
  translations: Record<string, string>,
  fontName = ""
): Promise<{ ready: boolean; fileName: string }> {
  return request(`/api/sessions/${sessionId}/apply`, {
    method: "POST",
    body: JSON.stringify({ translations, fontName }),
  });
}

export { ApiError };
