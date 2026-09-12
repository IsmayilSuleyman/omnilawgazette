import type { Metadata } from "next";
import { Inter, Fraunces, Jost } from "next/font/google";
import SiteHeader from "@/components/SiteHeader";
import "./globals.css";

const inter = Inter({
  variable: "--font-inter",
  subsets: ["latin"],
});

const fraunces = Fraunces({
  variable: "--font-fraunces",
  subsets: ["latin"],
});

const jost = Jost({
  variable: "--font-jost",
  subsets: ["latin"],
  weight: ["200", "300", "400", "500"],
});

export const metadata: Metadata = {
  title: {
    default: "Omni Law Gazette",
    template: "%s · Omni Law Gazette",
  },
  description:
    "Weekly legislature digest of Omni Law Firm — the internal library of weekly issues for Omni colleagues.",
  robots: { index: false, follow: false },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html
      lang="en"
      className={`${inter.variable} ${fraunces.variable} ${jost.variable} h-full antialiased`}
    >
      <body className="min-h-full flex flex-col">
        <div className="aurora" aria-hidden />
        <SiteHeader />
        <main className="flex-1 w-full max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">{children}</main>
        <footer className="mt-20 border-t hairline">
          <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 flex flex-col sm:flex-row gap-3 items-center justify-between text-xs text-silver">
            <p>
              © {new Date().getFullYear()} Omni Law Firm — Law Gazette. Internal circulation only.
            </p>
            <p className="font-jost tracking-[0.3em] uppercase text-[0.65rem]">
              omni · law firm
            </p>
          </div>
        </footer>
      </body>
    </html>
  );
}
