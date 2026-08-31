import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Agency Deck Translator",
  description: "AI 기반 프레젠테이션/카피 번역 도구",
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html lang="ko">
      <body>{children}</body>
    </html>
  );
}
