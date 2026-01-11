// WASMモジュールの型定義
// 実際のWASMモジュールがビルドされたら、このファイルを更新します

export interface TableData {
  headers: string[];
  rows: string[][];
}

export interface CompareOptions {
  trim: boolean;
  case_insensitive: boolean;
}

export interface CompareInput {
  left_headers: string[];
  left_rows: string[][];
  right_headers: string[];
  right_rows: string[][];
  key: string;
  options: CompareOptions;
}

export interface CompareOutput {
  result: TableData;
  left_only: TableData;
  right_only: TableData;
  duplicates: TableData;
  log: [string, string][];
}

export interface SplitInput {
  headers: string[];
  rows: string[][];
  key: string;
}

export interface SplitPart {
  key_value: string;
  table: TableData;
}

export interface SplitOutput {
  parts: SplitPart[];
}

// WASM関数の型定義
export interface WasmModule {
  compare_files(input_json: string): string;
  split_file(input_json: string): string;
}

// WASMモジュールをロードする関数
export async function loadWasmModule(): Promise<WasmModule | null> {
  try {
    // ビルドされたWASMモジュールをインポート
    // wasm-pack build --target web でビルド後、以下のパスでインポートできます
    // 本番環境では動的インポートを使用
    const wasm = process.env.NODE_ENV === 'production'
      ? await import('../excel-merge-wasm/pkg/excel_merge_wasm')
      : await import('../excel-merge-wasm/pkg/excel_merge_wasm');
    
    // 初期化関数を呼び出してWASMモジュールを初期化
    if (wasm.default) {
      await wasm.default();
    }
    return {
      compare_files: wasm.compare_files,
      split_file: wasm.split_file,
    };
  } catch (error) {
    // エラーログを出力（本番環境でも確認できるように）
    console.error('WASMモジュールのロードに失敗しました:', error);
    return null;
  }
}

