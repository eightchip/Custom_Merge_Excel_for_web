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
    return config;
  },
};

export default nextConfig;
