import "../styles/globals.css";

export const metadata = {
  title: "空余时间选择与导出",
  description: "体育馆空余时间查询与模板导出",
};

export default function RootLayout({ children }) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
