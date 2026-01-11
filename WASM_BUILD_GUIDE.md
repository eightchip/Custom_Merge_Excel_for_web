# WASMモジュールのビルドガイド

## 前提条件

1. **Rust** がインストールされている必要があります
   - インストール方法: https://www.rust-lang.org/tools/install
   - `rustc --version` で確認

2. **wasm-pack** がインストールされている必要があります
   - インストール方法:
     ```powershell
     cargo install wasm-pack
     ```
   - `wasm-pack --version` で確認

## ビルド手順

### 方法1: npmスクリプトを使用（推奨）

```powershell
cd custom-merge-excel-web
npm run build:wasm
```

### 方法2: 直接wasm-packを実行

```powershell
cd custom-merge-excel-web\excel-merge-wasm
wasm-pack build --target web --out-dir pkg
```

## ビルド後のファイル構造

ビルドが成功すると、以下のディレクトリとファイルが生成されます：

```
excel-merge-wasm/
  pkg/
    excel_merge_wasm.js       # JavaScriptバインディング
    excel_merge_wasm_bg.wasm  # WASMバイナリ
    excel_merge_wasm.d.ts     # TypeScript型定義
    package.json              # パッケージ情報
    ...その他のファイル
```

## トラブルシューティング

### wasm-packがインストールされていない場合

```powershell
cargo install wasm-pack
```

### Rustがインストールされていない場合

1. https://www.rust-lang.org/tools/install にアクセス
2. インストーラーをダウンロードして実行
3. インストール後、PowerShellを再起動

### ビルドエラーが発生する場合

```powershell
# クリーンビルド
cd custom-merge-excel-web\excel-merge-wasm
cargo clean
wasm-pack build --target web --out-dir pkg
```

## Next.jsでの使用

ビルド後、Next.jsアプリケーションでWASMモジュールが自動的に使用されます。

```typescript
// lib/wasm-types.ts で自動的にインポートされます
const wasm = await loadWasmModule();
if (wasm) {
  const result = wasm.compare_files(inputJson);
}
```

