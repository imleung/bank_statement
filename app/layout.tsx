import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Mạch Tiền — Chuẩn hóa sao kê ngân hàng",
  description: "Gom, chuẩn hóa, phân tích và xuất sao kê từ nhiều ngân hàng ngay trên thiết bị của bạn.",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="vi">
      <body>{children}</body>
    </html>
  );
}
