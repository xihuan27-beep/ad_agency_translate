export type Direction = "ko_en" | "en_ko";

export type Category = "presentation" | "copy";

export interface TextUnit {
  id: string;
  slide_idx: number;
  shape_id: number;
  p_idx: number;
  ko_text: string;
  font_size: number;
  shape_text: string;
  shape_para_count: number;
}

export interface ClassifiedUnit extends TextUnit {
  category: Category;
}

export interface KeyPhrasePair {
  한국어: string;
  영어: string;
}

export interface PresentationTranslation {
  en_text: string;
  notes: string;
  clarification: string;
}

export interface CopyOptionSet {
  options: [string, string, string];
  notes: string;
  recommendation: string;
  cultural_flag: string;
  clarification: string;
}

export type Stage = "upload" | "classify" | "review_2a" | "review_2b" | "en_ko" | "download";
