import { create } from "zustand";
import type { ClassifiedUnit, CopyOptionSet, Direction, KeyPhrasePair, PresentationTranslation, Stage, TextUnit } from "./types";

interface AppState {
  sessionId: string | null;
  stage: Stage;
  direction: Direction;

  // upload
  fetchedUrl: string;
  fileType: "pptx" | "docx";
  slideCount: number;
  hasSlideImages: boolean;
  keyPhrases: KeyPhrasePair[];

  // classify
  textUnits: TextUnit[];
  classifiedUnits: ClassifiedUnit[];
  excludedIds: Set<string>;
  mergedSeps: Record<number, Set<number>>; // slideIdx -> set of positions merged with next
  activeSlide: number;

  // review stages
  presentationUnits: TextUnit[];
  copyUnits: TextUnit[];
  presentationTranslations: Record<string, PresentationTranslation>;
  copyOptions: Record<string, CopyOptionSet>;
  copySelections: Record<string, string>;
  currentPresIdx: number;
  currentCopyIdx: number;

  // en_ko
  enKoTranslations: Record<string, string>;

  // download
  fileName: string;

  setSessionId: (id: string) => void;
  setStage: (stage: Stage) => void;
  setDirection: (dir: Direction) => void;
  setFetchedUrl: (url: string) => void;
  setKeyPhrases: (kp: KeyPhrasePair[]) => void;
  setUploadResult: (r: { fileType: "pptx" | "docx"; slideCount: number; hasSlideImages: boolean; textUnits: TextUnit[] }) => void;
  setClassifiedUnits: (units: ClassifiedUnit[]) => void;
  toggleCategory: (id: string) => void;
  setAllCategory: (slideIdx: number, category: "presentation" | "copy") => void;
  toggleExcluded: (id: string) => void;
  toggleMerge: (slideIdx: number, pos: number) => void;
  addManualUnit: (unit: ClassifiedUnit) => void;
  setActiveSlide: (idx: number) => void;
  setReviewUnits: (pres: TextUnit[], copy: TextUnit[]) => void;
  resetReviewProgress: () => void;
  setPresentationTranslations: (t: Record<string, PresentationTranslation>) => void;
  updatePresentationTranslation: (id: string, enText: string) => void;
  setCopyOptions: (o: Record<string, CopyOptionSet>) => void;
  setCopySelection: (id: string, text: string) => void;
  setCopySelectionsForIds: (ids: string[], text: string) => void;
  setCopySelectionsMap: (map: Record<string, string>) => void;
  updateCopyOptionText: (id: string, optionIdx: number, text: string) => void;
  setCurrentPresIdx: (i: number) => void;
  setCurrentCopyIdx: (i: number) => void;
  setEnKoTranslations: (t: Record<string, string>) => void;
  updateEnKoTranslation: (id: string, text: string) => void;
  setFileName: (n: string) => void;
  reset: () => void;
}

const initialState = {
  sessionId: null as string | null,
  stage: "upload" as Stage,
  direction: "ko_en" as Direction,
  fetchedUrl: "",
  fileType: "pptx" as "pptx" | "docx",
  slideCount: 0,
  hasSlideImages: false,
  keyPhrases: [{ 한국어: "", 영어: "" }] as KeyPhrasePair[],
  textUnits: [] as TextUnit[],
  classifiedUnits: [] as ClassifiedUnit[],
  excludedIds: new Set<string>(),
  mergedSeps: {} as Record<number, Set<number>>,
  activeSlide: 0,
  presentationUnits: [] as TextUnit[],
  copyUnits: [] as TextUnit[],
  presentationTranslations: {} as Record<string, PresentationTranslation>,
  copyOptions: {} as Record<string, CopyOptionSet>,
  copySelections: {} as Record<string, string>,
  currentPresIdx: 0,
  currentCopyIdx: 0,
  enKoTranslations: {} as Record<string, string>,
  fileName: "translated.pptx",
};

export const useAppStore = create<AppState>((set) => ({
  ...initialState,

  setSessionId: (id) => set({ sessionId: id }),
  setStage: (stage) => set({ stage }),
  setDirection: (direction) => set({ direction }),
  setFetchedUrl: (fetchedUrl) => set({ fetchedUrl }),
  setKeyPhrases: (keyPhrases) => set({ keyPhrases }),
  setUploadResult: (r) =>
    set({
      fileType: r.fileType,
      slideCount: r.slideCount,
      hasSlideImages: r.hasSlideImages,
      textUnits: r.textUnits,
    }),
  setClassifiedUnits: (classifiedUnits) => set({ classifiedUnits }),

  toggleCategory: (id) =>
    set((s) => ({
      classifiedUnits: s.classifiedUnits.map((u) =>
        u.id === id ? { ...u, category: u.category === "presentation" ? "copy" : "presentation" } : u
      ),
    })),

  setAllCategory: (slideIdx, category) =>
    set((s) => ({
      classifiedUnits: s.classifiedUnits.map((u) => (u.slide_idx === slideIdx ? { ...u, category } : u)),
    })),

  toggleExcluded: (id) =>
    set((s) => {
      const next = new Set(s.excludedIds);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return { excludedIds: next };
    }),

  toggleMerge: (slideIdx, pos) =>
    set((s) => {
      const next = { ...s.mergedSeps };
      const set_ = new Set(next[slideIdx] || []);
      if (set_.has(pos)) set_.delete(pos);
      else set_.add(pos);
      next[slideIdx] = set_;
      return { mergedSeps: next };
    }),

  addManualUnit: (unit) => set((s) => ({ classifiedUnits: [...s.classifiedUnits, unit] })),

  setActiveSlide: (activeSlide) => set({ activeSlide }),

  setReviewUnits: (presentationUnits, copyUnits) => set({ presentationUnits, copyUnits }),
  resetReviewProgress: () =>
    set({
      presentationTranslations: {},
      copyOptions: {},
      copySelections: {},
      currentPresIdx: 0,
      currentCopyIdx: 0,
    }),
  setPresentationTranslations: (presentationTranslations) => set({ presentationTranslations }),
  updatePresentationTranslation: (id, enText) =>
    set((s) => ({
      presentationTranslations: {
        ...s.presentationTranslations,
        [id]: { ...s.presentationTranslations[id], en_text: enText },
      },
    })),

  setCopyOptions: (copyOptions) => set({ copyOptions }),
  setCopySelection: (id, text) => set((s) => ({ copySelections: { ...s.copySelections, [id]: text } })),
  setCopySelectionsForIds: (ids, text) =>
    set((s) => {
      const next = { ...s.copySelections };
      for (const id of ids) next[id] = text;
      return { copySelections: next };
    }),
  setCopySelectionsMap: (map) => set((s) => ({ copySelections: { ...s.copySelections, ...map } })),
  updateCopyOptionText: (id, optionIdx, text) =>
    set((s) => {
      const current = s.copyOptions[id];
      if (!current) return {};
      const options = [...current.options] as [string, string, string];
      const wasSelected = s.copySelections[id] === options[optionIdx];
      options[optionIdx] = text;
      const nextSelections = wasSelected ? { ...s.copySelections, [id]: text } : s.copySelections;
      return {
        copyOptions: { ...s.copyOptions, [id]: { ...current, options } },
        copySelections: nextSelections,
      };
    }),

  setCurrentPresIdx: (currentPresIdx) => set({ currentPresIdx }),
  setCurrentCopyIdx: (currentCopyIdx) => set({ currentCopyIdx }),
  setEnKoTranslations: (enKoTranslations) => set({ enKoTranslations }),
  updateEnKoTranslation: (id, text) =>
    set((s) => ({ enKoTranslations: { ...s.enKoTranslations, [id]: text } })),
  setFileName: (fileName) => set({ fileName }),

  reset: () => set({ ...initialState, excludedIds: new Set(), mergedSeps: {} }),
}));

export function materializeMerges(units: ClassifiedUnit[], mergedSeps: Record<number, Set<number>>): ClassifiedUnit[] {
  const bySlide = new Map<number, ClassifiedUnit[]>();
  for (const u of units) {
    const arr = bySlide.get(u.slide_idx) || [];
    arr.push(u);
    bySlide.set(u.slide_idx, arr);
  }
  const result: ClassifiedUnit[] = [];
  for (const [slideIdx, items] of bySlide) {
    const seps = mergedSeps[slideIdx] || new Set<number>();
    let pos = 0;
    const n = items.length;
    while (pos < n) {
      const run = [items[pos]];
      while (seps.has(pos) && pos + 1 < n) {
        pos += 1;
        run.push(items[pos]);
      }
      if (run.length === 1) {
        result.push(run[0]);
      } else {
        const merged: ClassifiedUnit = {
          ...run[0],
          ko_text: run.map((r) => r.ko_text).join(" "),
          shape_text: run.map((r) => r.ko_text).join(" "),
        };
        result.push(merged);
      }
      pos += 1;
    }
  }
  return result;
}
