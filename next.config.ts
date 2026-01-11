import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Webpack設定（本番ビルドで使用）
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
