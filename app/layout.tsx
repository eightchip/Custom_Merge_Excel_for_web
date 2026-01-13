import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";
import LayoutClient from "./layout-client";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: "Custom Merge Excel Web",
  description: "高速Excelファイル統合・分割ツール",
  manifest: "/manifest.json",
  appleWebApp: {
    capable: true,
    statusBarStyle: "default",
    title: "Merge Excel",
  },
};

export const viewport = {
  width: "device-width",
  initialScale: 1,
  maximumScale: 1,
  themeColor: "#000000",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" className="dark" suppressHydrationWarning>
      <body
        className={`${geistSans.variable} ${geistMono.variable} antialiased`}
      >
        <script
          dangerouslySetInnerHTML={{
            __html: `
              (function() {
                try {
                  const savedTheme = localStorage.getItem('theme');
                  const root = document.documentElement;
                  root.className = root.className.replace(/theme-\\w+/g, '').replace(/\\bdark\\b/g, '');
                  if (savedTheme === 'dark') {
                    root.classList.add('dark');
                  } else if (savedTheme === 'ocean') {
                    root.classList.add('theme-ocean');
                  } else if (savedTheme === 'forest') {
                    root.classList.add('theme-forest');
                  } else {
                    root.classList.add('dark');
                  }
                } catch (e) {
                  document.documentElement.classList.add('dark');
                }
              })();
            `,
          }}
        />
        <LayoutClient />
        {children}
      </body>
    </html>
  );
}
