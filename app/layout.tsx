import type { Metadata } from "next";
import "./styles.css";

export const metadata: Metadata = {
  title: "CPE Scheduling App",
  description: "Online console for creating exam dates, unavailability sheets, and duty schedules."
};

export default function RootLayout({
  children
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
