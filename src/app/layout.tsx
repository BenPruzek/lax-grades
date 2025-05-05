import type { Metadata } from "next";
import { Inter } from "next/font/google";
import "./globals.css";
import Footer from "@/components/component/shared-components/footer";
import NavBar from "@/components/component/shared-components/navbar";
import Script from "next/script";

const inter = Inter({ subsets: ["latin"] });

export const metadata: Metadata = {
  title: "LAXGRADES - Explore UWL Grade Distributions",
  description: "Discover grade distributions for classes at the University of Wisconsin-La Crosse. Search and compare historical grade data, GPAs, and pass rates for courses across various semesters and instructors.",
  keywords: "UWL, University of Wisconsin-La Crosse, grade distributions, GPA, pass rates, courses, classes, instructors, semesters",
};


export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <head>
        {/* AdSense Verification Meta Tag */}
        <meta name="google-adsense-account" content="ca-pub-1116327136347209"></meta>

        {/* AdSense Script */}
        <script async src="https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=ca-pub-1116327136347209"
              crossOrigin="anonymous"></script>

        {/* Google Analytics Scripts */}
        <Script async src="https://www.googletagmanager.com/gtag/js?id=G-XYGL070WW8"></Script>
        <Script
          id="gtm-script"
          dangerouslySetInnerHTML={{
            __html: `
              window.dataLayer = window.dataLayer || [];
              function gtag(){dataLayer.push(arguments);}
              gtag('js', new Date());
              
              gtag('config', 'G-XYGL070WW8');
              `
          }}
        />
      </head>
      <body className={inter.className}>
        <NavBar />
        {children}
        <Footer />
      </body>
    </html>
  );
}
