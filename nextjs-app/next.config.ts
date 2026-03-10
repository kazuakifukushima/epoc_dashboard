import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // 静的エクスポート: BUILD_STATIC=1 npm run build
  ...(process.env.BUILD_STATIC === "1" ? { output: "export" as const } : {}),
};

export default nextConfig;
