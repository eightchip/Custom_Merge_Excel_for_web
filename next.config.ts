import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Turbopack用の設定（WASMは自動的にサポートされます）
  turbopack: {},
  // Webpack用の設定（--webpackフラグ使用時）
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
  swcMinify: true,
};

export default nextConfig;
