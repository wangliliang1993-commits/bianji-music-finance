import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "边际音乐 · 财务经营管理",
  description: "琴行财务与进销存管理系统",
  manifest: "/manifest.webmanifest",
  themeColor: "#2b82ac"
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
