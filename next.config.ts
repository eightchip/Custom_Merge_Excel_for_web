import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Turbopack設定（Next.js 16のデフォルト、空の設定でエラーを回避）
  turbopack: {},
  // Webpack設定（本番ビルドで使用、WASMサポートのため必要）
  webpack: (config, { isServer }) => {
    if (!isServer) {
      config.experiments = {
        ...config.experiments,
        asyncWebAssembly: true,
      };
    }
    // WASMファイルを適切に処理
    config.resolve.fallback = {
      ...config.resolve.fallback,
      fs: false,
    };
    return config;
  },
  // 静的アセットの最適化
  poweredByHeader: false,
  // 圧縮設定
  compress: true,
};

export default nextConfig;
