"use client";

import { useState, useRef, useMemo, useEffect } from "react";
import { Tabs, TabsList, TabsTrigger, TabsContent } from "@/components/ui/tabs";
import { Button } from "@/components/ui/button";
import { Card, CardHeader, CardTitle, CardDescription, CardContent } from "@/components/ui/card";
import { Checkbox } from "@/components/ui/checkbox";
import { Upload, FileSpreadsheet, Download, X, ChevronUp, ChevronDown, Sliders } from "lucide-react";
import { readExcelFile, writeExcelFile, type TableData } from "@/lib/excel-utils";
import { loadWasmModule, type CompareOptions, type CompareInput } from "@/lib/wasm-types";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import JSZip from "jszip";

type Theme = "light" | "dark" | "ocean" | "forest";

const themes: { id: Theme; name: string; description: string }[] = [
  { id: "light", name: "ライト", description: "白背景・黒文字" },
  { id: "dark", name: "ダーク", description: "黒背景・白文字" },
  { id: "ocean", name: "オーシャン", description: "青緑系" },
  { id: "forest", name: "フォレスト", description: "緑系" },
];

// プレビューテーブルコンポーネント
function PreviewTable({ data }: { data: TableData }) {
  if (!data || data.rows.length === 0) {
    return <div className="text-sm text-muted-foreground">データがありません</div>;
  }

  return (
    <div className="max-h-96 overflow-auto rounded-md border">
      <table className="w-full text-sm">
        <thead className="bg-muted sticky top-0">
          <tr>
            {data.headers.map((header, idx) => (
              <th key={idx} className="px-4 py-2 text-left font-medium border-b">
                {header}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {data.rows.slice(0, 100).map((row, rowIdx) => (
            <tr key={rowIdx} className="border-b hover:bg-muted/50">
              {row.map((cell, cellIdx) => (
                <td key={cellIdx} className="px-4 py-2 max-w-xs truncate" title={cell}>
                  {cell}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
      {data.rows.length > 100 && (
        <div className="p-2 text-xs text-muted-foreground text-center border-t">
          最初の100行を表示しています（全{data.rows.length}行）
        </div>
      )}
    </div>
  );
}

export default function Home() {
  const [activeTab, setActiveTab] = useState("compare");
  const [currentTheme, setCurrentTheme] = useState<Theme>("dark");
  const [showThemeMenu, setShowThemeMenu] = useState(false);

  // Compare state
  const [leftFile, setLeftFile] = useState<File | null>(null);
  const [rightFile, setRightFile] = useState<File | null>(null);
  const [leftData, setLeftData] = useState<TableData | null>(null);
  const [rightData, setRightData] = useState<TableData | null>(null);
  const [compareKeys, setCompareKeys] = useState<string[]>([]);
  const [compareColumns, setCompareColumns] = useState<{ left: string; right: string; label: string }[]>([]);
  const [compareOptions, setCompareOptions] = useState<CompareOptions>({
    trim: true,
    case_insensitive: false,
  });
  const [sortByKeys, setSortByKeys] = useState(true); // キー列でソートするかどうか
  const [excelHeaderColor, setExcelHeaderColor] = useState(true); // ヘッダー行に色を付ける
  const [excelBorders, setExcelBorders] = useState(true); // 罫線を引く
  const [excelHeaderColorValue, setExcelHeaderColorValue] = useState("aqua"); // ヘッダー行の色
  const [excelShowTotal, setExcelShowTotal] = useState(true); // 合計行を表示する

  // Excelヘッダー行のカラーパレット（視認性の良い10種類）
  const excelHeaderColors = [
    { id: "aqua", name: "アクア", color: "#B3E5FC", argb: "FFB3E5FC" }, // Accent5 lighter80%
    { id: "blue", name: "青", color: "#90CAF9", argb: "FF90CAF9" },
    { id: "green", name: "グリーン", color: "#A5D6A7", argb: "FFA5D6A7" },
    { id: "orange", name: "オレンジ", color: "#FFCC80", argb: "FFFFCC80" },
    { id: "purple", name: "パープル", color: "#CE93D8", argb: "FFCE93D8" },
    { id: "pink", name: "ピンク", color: "#F48FB1", argb: "FFF48FB1" },
    { id: "yellow", name: "イエロー", color: "#FFF59D", argb: "FFFFF59D" },
    { id: "teal", name: "ティール", color: "#80CBC4", argb: "FF80CBC4" },
    { id: "cyan", name: "シアン", color: "#80DEEA", argb: "FF80DEEA" },
    { id: "lime", name: "ライム", color: "#E6EE9C", argb: "FFE6EE9C" },
  ];
  const [compareResult, setCompareResult] = useState<any | null>(null);
  const [mergedResult, setMergedResult] = useState<TableData | null>(null);
  const [selectedColumns, setSelectedColumns] = useState<string[]>([]);
  const [sortColumns, setSortColumns] = useState<{ column: string; direction: "asc" | "desc" }[]>([]);
  
  // 選択された列のみを含むテーブルデータを生成
  const filterColumns = (data: TableData, columns: string[]): TableData => {
    const columnIndices = columns.map(col => data.headers.indexOf(col)).filter(idx => idx !== -1);
    return {
      headers: columns.filter(col => data.headers.includes(col)),
      rows: data.rows.map(row => columnIndices.map(idx => row[idx] || "")),
    };
  };
  
  // リアルタイムソート処理（Hooksの順序を保つため、条件分岐の外に配置）
  const sortedMergedResult = useMemo(() => {
    if (!mergedResult) return null;
    
    // 選択された列でフィルタリング
    const filteredData = filterColumns(mergedResult, selectedColumns);
    
    if (sortColumns.length === 0) return filteredData;
    
    const sortedRows = [...filteredData.rows].sort((a, b) => {
      for (const sortCol of sortColumns) {
        const sortIdx = filteredData.headers.indexOf(sortCol.column);
        if (sortIdx === -1) continue;
        
        const aVal = parseFloat(a[sortIdx]) || 0;
        const bVal = parseFloat(b[sortIdx]) || 0;
        // 文字列の場合は文字列比較
        if (isNaN(aVal) || isNaN(bVal)) {
          const aStr = String(a[sortIdx] || "");
          const bStr = String(b[sortIdx] || "");
          const result = sortCol.direction === "asc" ? aStr.localeCompare(bStr) : bStr.localeCompare(aStr);
          if (result !== 0) return result;
        } else {
          const result = sortCol.direction === "asc" ? aVal - bVal : bVal - aVal;
          if (result !== 0) return result;
        }
      }
      return 0;
    });
    
    return {
      ...filteredData,
      rows: sortedRows,
    };
  }, [mergedResult, selectedColumns, sortColumns]);

  // Split state
  const [splitFile, setSplitFile] = useState<File | null>(null);
  const [splitData, setSplitData] = useState<TableData | null>(null);
  const [splitKeys, setSplitKeys] = useState<string[]>([]);
  const [splitResult, setSplitResult] = useState<any | null>(null);
  const [selectedSplitColumns, setSelectedSplitColumns] = useState<string[]>([]);
  const [splitNumericColumns, setSplitNumericColumns] = useState<string[]>([]);
  const [splitSortColumns, setSplitSortColumns] = useState<{ column: string; direction: "asc" | "desc" }[]>([]);
  
  // 分割モードのソート処理（Hooksの順序を保つため、条件分岐の外に配置）
  const sortedSplitPreviewData = useMemo(() => {
    if (!splitResult || splitResult.parts.length === 0) return null;
    
    // 最初のファイルのデータをソート（選択された列でフィルタリング済み）
    const firstPart = splitResult.parts[0];
    const filteredData = filterColumns(firstPart.table, selectedSplitColumns);
    
    if (splitSortColumns.length === 0) return filteredData;
    
    const sortedRows = [...filteredData.rows].sort((a, b) => {
      for (const sortCol of splitSortColumns) {
        const sortIdx = filteredData.headers.indexOf(sortCol.column);
        if (sortIdx === -1) continue;
        
        const aVal = parseFloat(a[sortIdx]) || 0;
        const bVal = parseFloat(b[sortIdx]) || 0;
        if (isNaN(aVal) || isNaN(bVal)) {
          const aStr = String(a[sortIdx] || "");
          const bStr = String(b[sortIdx] || "");
          const result = sortCol.direction === "asc" ? aStr.localeCompare(bStr) : bStr.localeCompare(aStr);
          if (result !== 0) return result;
        } else {
          const result = sortCol.direction === "asc" ? aVal - bVal : bVal - aVal;
          if (result !== 0) return result;
        }
      }
      return 0;
    });
    
    return {
      ...filteredData,
      rows: sortedRows,
    };
  }, [splitResult, selectedSplitColumns, splitSortColumns]);

  const leftFileInputRef = useRef<HTMLInputElement>(null);
  const rightFileInputRef = useRef<HTMLInputElement>(null);
  const splitFileInputRef = useRef<HTMLInputElement>(null);
  const themeMenuRef = useRef<HTMLDivElement>(null);

  // テーマ切り替え
  useEffect(() => {
    const root = document.documentElement;
    // 既存のテーマクラスを削除
    root.className = root.className.replace(/theme-\w+/g, "").replace(/\bdark\b/g, "");
    if (currentTheme === "dark") {
      root.classList.add("dark");
    } else if (currentTheme !== "light") {
      root.classList.add(`theme-${currentTheme}`);
    }
    // ローカルストレージに保存
    localStorage.setItem("theme", currentTheme);
  }, [currentTheme]);

  // 初期テーマ読み込み
  useEffect(() => {
    const savedTheme = localStorage.getItem("theme") as Theme | null;
    if (savedTheme && themes.some(t => t.id === savedTheme)) {
      setCurrentTheme(savedTheme);
    } else {
      // 保存されたテーマが無効な場合、デフォルト（ダーク）を設定
      setCurrentTheme("dark");
    }
  }, []);

  // テーマメニューの外側クリックで閉じる
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (themeMenuRef.current && !themeMenuRef.current.contains(event.target as Node)) {
        setShowThemeMenu(false);
      }
    };
    if (showThemeMenu) {
      document.addEventListener("mousedown", handleClickOutside);
      return () => document.removeEventListener("mousedown", handleClickOutside);
    }
  }, [showThemeMenu]);

  const handleLeftFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLeftFile(file);
    setLeftData(null); // リセット
    try {
      const data = await readExcelFile(file);
      console.log('左側ファイル読み込み成功:', data.headers.length, '列', data.rows.length, '行');
      setLeftData(data);
      if (compareKeys.length === 0 && data.headers.length > 0) {
        setCompareKeys([data.headers[0]]);
      }
    } catch (error) {
      console.error('左側ファイル読み込みエラー:', error);
      alert(`ファイルの読み込みに失敗しました: ${error}`);
      setLeftFile(null);
    }
  };

  const handleRightFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setRightFile(file);
    setRightData(null); // リセット
    try {
      const data = await readExcelFile(file);
      console.log('右側ファイル読み込み成功:', data.headers.length, '列', data.rows.length, '行');
      setRightData(data);
    } catch (error) {
      console.error('右側ファイル読み込みエラー:', error);
      alert(`ファイルの読み込みに失敗しました: ${error}`);
      setRightFile(null);
    }
  };

  const handleSplitFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setSplitFile(file);
    try {
      const data = await readExcelFile(file);
      setSplitData(data);
      if (splitKeys.length === 0 && data.headers.length > 0) {
        setSplitKeys([data.headers[0]]);
      }
    } catch (error) {
      alert(`ファイルの読み込みに失敗しました: ${error}`);
    }
  };

  // 複数キーを結合する関数
  const combineKeys = (row: string[], keyIndices: number[], options: CompareOptions): string => {
    const values = keyIndices.map(idx => row[idx] || "");
    let combined = values.join("|");
    if (options.trim) {
      combined = combined.trim();
    }
    if (options.case_insensitive) {
      combined = combined.toLowerCase();
    }
    return combined;
  };

  // キー列でソートする関数
  const sortByKeyColumns = (data: TableData, keyColumns: string[], options: CompareOptions): TableData => {
    const keyIndices = keyColumns.map(key => data.headers.indexOf(key)).filter(idx => idx !== -1);
    if (keyIndices.length === 0) return data;

    const sortedRows = [...data.rows].sort((a, b) => {
      for (const keyIdx of keyIndices) {
        let aVal = a[keyIdx] || "";
        let bVal = b[keyIdx] || "";
        
        if (options.trim) {
          aVal = aVal.trim();
          bVal = bVal.trim();
        }
        if (options.case_insensitive) {
          aVal = aVal.toLowerCase();
          bVal = bVal.toLowerCase();
        }
        
        // 数値として比較を試みる
        const aNum = parseFloat(aVal);
        const bNum = parseFloat(bVal);
        if (!isNaN(aNum) && !isNaN(bNum)) {
          const diff = aNum - bNum;
          if (diff !== 0) return diff;
        } else {
          // 文字列として比較
          const diff = aVal.localeCompare(bVal, 'ja');
          if (diff !== 0) return diff;
        }
      }
      return 0;
    });

    return {
      ...data,
      rows: sortedRows,
    };
  };

  const handleCompare = async () => {
    if (!leftData || !rightData || compareKeys.length === 0) {
      alert("両方のファイルとキー列を選択してください");
      return;
    }

    const wasm = await loadWasmModule();
    if (!wasm) {
      alert("WASMモジュールが利用できません。後でビルドしてください。");
      return;
    }

    try {
      // キー列でソートする場合、事前にソート
      let sortedLeftData = leftData;
      let sortedRightData = rightData;
      if (sortByKeys) {
        sortedLeftData = sortByKeyColumns(leftData, compareKeys, compareOptions);
        sortedRightData = sortByKeyColumns(rightData, compareKeys, compareOptions);
      }
      
      // 複数キーの場合、一時的に結合キー列を作成
      const combinedKeyName = compareKeys.join("|");
      
      // 左側のデータに結合キー列を追加
      const leftKeyIndices = compareKeys.map(key => sortedLeftData.headers.indexOf(key));
      const leftHeadersWithKey = [...sortedLeftData.headers, combinedKeyName];
      const leftRowsWithKey = sortedLeftData.rows.map(row => [
        ...row,
        combineKeys(row, leftKeyIndices, compareOptions)
      ]);

      // 右側のデータに結合キー列を追加
      const rightKeyIndices = compareKeys.map(key => sortedRightData.headers.indexOf(key));
      const rightHeadersWithKey = [...sortedRightData.headers, combinedKeyName];
      const rightRowsWithKey = sortedRightData.rows.map(row => [
        ...row,
        combineKeys(row, rightKeyIndices, compareOptions)
      ]);

      const input: CompareInput = {
        left_headers: leftHeadersWithKey,
        left_rows: leftRowsWithKey,
        right_headers: rightHeadersWithKey,
        right_rows: rightRowsWithKey,
        key: combinedKeyName,
        options: compareOptions,
      };

      const resultJson = wasm.compare_files(JSON.stringify(input));
      const result = JSON.parse(resultJson);
      
      // 結合キー列を結果から削除
      const removeCombinedKey = (data: TableData) => {
        const keyIdx = data.headers.indexOf(combinedKeyName);
        if (keyIdx === -1) return data;
        return {
          headers: data.headers.filter((_, i) => i !== keyIdx),
          rows: data.rows.map(row => row.filter((_, i) => i !== keyIdx))
        };
      };

      result.result = removeCombinedKey(result.result);
      result.left_only = removeCombinedKey(result.left_only);
      result.right_only = removeCombinedKey(result.right_only);
      result.duplicates = removeCombinedKey(result.duplicates);
      
      // マージ結果を生成（すべての行を含む）
      let mergedHeaders = result.result.headers;
      let mergedRows: string[][] = [
        ...result.result.rows,
        ...result.left_only.rows,
        ...result.right_only.rows,
        ...result.duplicates.rows,
      ];
      
      // 結合キー列を統合（L__とR__を1つの列に）
      const keyColumnMapping: Map<string, string> = new Map();
      const unifiedHeaders: string[] = [];
      const unifiedRows: string[][] = [];
      
      // ヘッダーを処理
      const processedKeys = new Set<string>();
      for (const header of mergedHeaders) {
        // L__またはR__で始まる結合キー列を検出
        const isKeyColumn = compareKeys.some(key => {
          return header === `L__${key}` || header === `R__${key}`;
        });
        
        if (isKeyColumn) {
          const keyName = header.replace(/^(L__|R__)/, '');
          if (!processedKeys.has(keyName)) {
            unifiedHeaders.push(keyName);
            keyColumnMapping.set(`L__${keyName}`, keyName);
            keyColumnMapping.set(`R__${keyName}`, keyName);
            processedKeys.add(keyName);
          }
        } else {
          unifiedHeaders.push(header);
        }
      }
      
      // 行を処理
      for (const row of mergedRows) {
        const unifiedRow: string[] = [];
        const rowMap = new Map<string, string>();
        
        // 各行の値をヘッダー名でマップ
        mergedHeaders.forEach((header: string, idx: number) => {
          rowMap.set(header, row[idx] || "");
        });
        
        // 統合されたヘッダー順に値を取得
        for (const header of unifiedHeaders) {
          // 結合キー列の場合、L__またはR__から値を取得（どちらかが存在すればその値を使用）
          if (processedKeys.has(header)) {
            const value = rowMap.get(`L__${header}`) || rowMap.get(`R__${header}`) || "";
            unifiedRow.push(value);
          } else {
            unifiedRow.push(rowMap.get(header) || "");
          }
        }
        
        unifiedRows.push(unifiedRow);
      }
      
      // 統合前の行マップを作成（差額計算用）
      const originalRowMaps = mergedRows.map((row) => {
        const map = new Map<string, string>();
        mergedHeaders.forEach((header: string, idx: number) => {
          map.set(header, row[idx] || "");
        });
        return map;
      });
      
      // 比較列の差額を計算して追加
      let finalHeaders = [...unifiedHeaders];
      let finalRows = unifiedRows.map((row, rowIdx) => [...row]);
      
      // 比較列の差額を計算
      compareColumns.forEach(col => {
        if (col.left && col.right && col.label) {
          const leftHeader = `L__${col.left}`;
          const rightHeader = `R__${col.right}`;
          const diffColumnName = col.label || `${col.left}-${col.right}`;
          
          // ヘッダーに差額列を追加
          finalHeaders.push(diffColumnName);
          
          // 各行で差額を計算（統合前の行マップを使用）
          finalRows = finalRows.map((row, rowIdx) => {
            const rowMap = originalRowMaps[rowIdx];
            const leftValue = parseFloat(rowMap.get(leftHeader) || "0") || 0;
            const rightValue = parseFloat(rowMap.get(rightHeader) || "0") || 0;
            const diff = leftValue - rightValue;
            return [...row, diff.toString()];
          });
        }
      });
      
      // ソート処理は後でリアルタイムで行うため、ここではソートしない
      const merged: TableData = {
        headers: finalHeaders,
        rows: finalRows,
      };
      setMergedResult(merged);
      
      // デフォルトで必須列（結合キー列）のみを選択
      const requiredColumns = finalHeaders.filter(header => compareKeys.includes(header));
      setSelectedColumns(requiredColumns);
      
      setCompareResult(result);
    } catch (error) {
      alert(`比較処理に失敗しました: ${error}`);
    }
  };

  const handleSplit = async () => {
    if (!splitData || splitKeys.length === 0) {
      alert("ファイルとキー列を選択してください");
      return;
    }

    const wasm = await loadWasmModule();
    if (!wasm) {
      alert("WASMモジュールが利用できません。後でビルドしてください。");
      return;
    }

    try {
      // 複数キーの場合、一時的に結合キー列を作成
      const combinedKeyName = splitKeys.join("|");
      
      // データに結合キー列を追加
      const keyIndices = splitKeys.map(key => splitData.headers.indexOf(key));
      const headersWithKey = [...splitData.headers, combinedKeyName];
      const rowsWithKey = splitData.rows.map(row => [
        ...row,
        combineKeys(row, keyIndices, { trim: true, case_insensitive: false })
      ]);

      const input = {
        headers: headersWithKey,
        rows: rowsWithKey,
        key: combinedKeyName,
      };

      const resultJson = wasm.split_file(JSON.stringify(input));
      const result = JSON.parse(resultJson);
      
      // 結合キー列を結果から削除
      const keyIdx = headersWithKey.indexOf(combinedKeyName);
      result.parts = result.parts.map((part: any) => ({
        ...part,
        table: {
          headers: part.table.headers.filter((_: any, i: number) => i !== keyIdx),
          rows: part.table.rows.map((row: string[]) => row.filter((_: string, i: number) => i !== keyIdx)),
        },
      }));
      
      // デフォルトでキー列のみを選択
      if (result.parts.length > 0) {
        const allHeaders = result.parts[0].table.headers;
        const requiredColumns = allHeaders.filter((header: string) => splitKeys.includes(header));
        setSelectedSplitColumns(requiredColumns);
      }
      
      setSplitResult(result);
    } catch (error) {
      alert(`分割処理に失敗しました: ${error}`);
    }
  };

  const handleDownloadCompare = async () => {
    if (!compareResult || !mergedResult) return;

    // ソート済みの結果を使用（sortedMergedResultは既に選択列でフィルタリング済み、ソート済み）
    const filteredMerged = sortedMergedResult || filterColumns(mergedResult, selectedColumns);

    // 金額列（比較列、差額列）を識別
    const amountColumnHeaders = new Set<string>();
    compareColumns.forEach(col => {
      if (col.left && col.right) {
        amountColumnHeaders.add(`L__${col.left}`);
        amountColumnHeaders.add(`R__${col.right}`);
        if (col.label) {
          amountColumnHeaders.add(col.label);
        }
      }
    });
    
    // 金額列のインデックスを取得（フィルタ後のヘッダーで）
    const amountColumnIndices = filteredMerged.headers
      .map((header, idx) => {
        // 比較列や差額列を検出
        const isAmountColumn = amountColumnHeaders.has(header) || 
                               compareColumns.some(col => col.label === header) ||
                               (header.includes('L__') && (header.includes('残高') || header.includes('借方') || header.includes('貸方') || header.includes('金額') || header.includes('発生')));
        return isAmountColumn ? idx : -1;
      })
      .filter(idx => idx !== -1);

    // 各金額列で小数点以下があるかチェック
    const hasDecimalPlaces = amountColumnIndices.map(colIdx => {
      return filteredMerged.rows.some(row => {
        const val = row[colIdx] || "";
        const num = parseFloat(val);
        if (!isNaN(num)) {
          return num % 1 !== 0; // 小数点以下があるか
        }
        return false;
      });
    });

    // 合計行を計算
    const totals: (string | number)[] = filteredMerged.headers.map((header, idx) => {
      if (amountColumnIndices.includes(idx)) {
        const sum = filteredMerged.rows.reduce((acc, row) => {
          const val = parseFloat(row[idx] || "0") || 0;
          return acc + val;
        }, 0);
        return sum;
      }
      return idx === 0 ? "合計" : "";
    });

    // ExcelJSを使用してExcelファイルを作成
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Sheet1");

    // ヘッダー行を追加
    worksheet.addRow(filteredMerged.headers);

    // データ行を追加
    filteredMerged.rows.forEach(row => {
      const rowData = row.map((cell, idx) => {
        if (amountColumnIndices.includes(idx)) {
          const num = parseFloat(cell || "0");
          return isNaN(num) ? cell : num;
        }
        return cell;
      });
      worksheet.addRow(rowData);
    });

    // 合計行を追加（オプション）
    if (excelShowTotal) {
      worksheet.addRow(totals);
    }

    // スタイルを適用
    const selectedColor = excelHeaderColors.find(c => c.id === excelHeaderColorValue) || excelHeaderColors[0];
    const totalRowNumber = excelShowTotal ? filteredMerged.rows.length + 2 : filteredMerged.rows.length + 1;
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        // ヘッダー行（1行目）に色を付ける
        if (rowNumber === 1 && excelHeaderColor) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: selectedColor.argb }
          };
          cell.font = {
            bold: true,
            color: { argb: 'FF000000' }
          };
        }

        // 罫線を引く
        if (excelBorders) {
          cell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
          };
        }

        // 金額列の数値形式を設定（小数点以下がある場合は表示）
        if (amountColumnIndices.includes(colNumber - 1)) {
          const amountIdx = amountColumnIndices.indexOf(colNumber - 1);
          const hasDecimal = hasDecimalPlaces[amountIdx];
          
          if (rowNumber > 1 && rowNumber <= totalRowNumber) {
            // データ行と合計行
            cell.numFmt = hasDecimal ? '#,##0.00' : '#,##0';
          }
        }
      });
    });

    // ファイルをダウンロード
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "merged_result.xlsx";
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleDownloadSplit = async () => {
    if (!splitResult) return;

    const zip = new JSZip();
    for (const part of splitResult.parts) {
      // ファイル名に使用できない文字を置換
      const safeFileName = part.key_value
        .replace(/[<>:"/\\|?*]/g, "_")
        .replace(/\s+/g, "_");
      
      // 選択された列のみを含むデータを生成
      let filteredData = filterColumns(part.table, selectedSplitColumns);
      
      // ソート処理（3列まで順位指定）
      if (splitSortColumns.length > 0) {
        const sortedRows = [...filteredData.rows].sort((a, b) => {
          for (const sortCol of splitSortColumns) {
            const sortIdx = filteredData.headers.indexOf(sortCol.column);
            if (sortIdx === -1) continue;
            
            const aVal = parseFloat(a[sortIdx]) || 0;
            const bVal = parseFloat(b[sortIdx]) || 0;
            if (isNaN(aVal) || isNaN(bVal)) {
              const aStr = String(a[sortIdx] || "");
              const bStr = String(b[sortIdx] || "");
              const result = sortCol.direction === "asc" ? aStr.localeCompare(bStr) : bStr.localeCompare(aStr);
              if (result !== 0) return result;
            } else {
              const result = sortCol.direction === "asc" ? aVal - bVal : bVal - aVal;
              if (result !== 0) return result;
            }
          }
          return 0;
        });
        filteredData = {
          ...filteredData,
          rows: sortedRows,
        };
      }
      
      // ExcelJSを使用してExcelファイルを作成
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Sheet1");

      // ヘッダー行を追加
      worksheet.addRow(filteredData.headers);

      // データ行を追加
      filteredData.rows.forEach(row => {
        const rowData = row.map((cell, cellIdx) => {
          const header = filteredData.headers[cellIdx];
          if (splitNumericColumns.includes(header)) {
            const num = parseFloat(cell || "0");
            return isNaN(num) ? cell : num;
          }
          return cell;
        });
        worksheet.addRow(rowData);
      });

      // 合計行を追加（数値列のみ）
      const numericColumnIndices = filteredData.headers
        .map((header, idx) => splitNumericColumns.includes(header) ? idx : -1)
        .filter(idx => idx !== -1);

      // 各数値列で小数点以下があるかチェック
      const hasDecimalPlaces = numericColumnIndices.map(colIdx => {
        return filteredData.rows.some(row => {
          const val = row[colIdx] || "";
          const num = parseFloat(val);
          if (!isNaN(num)) {
            return num % 1 !== 0; // 小数点以下があるか
          }
          return false;
        });
      });

      if (numericColumnIndices.length > 0 && excelShowTotal) {
        const totals: (string | number)[] = filteredData.headers.map((header, idx) => {
          if (splitNumericColumns.includes(header)) {
            const sum = filteredData.rows.reduce((acc, row) => {
              const val = parseFloat(row[idx] || "0") || 0;
              return acc + val;
            }, 0);
            return sum;
          }
          return idx === 0 ? "合計" : "";
        });
        worksheet.addRow(totals);
      }

      // スタイルを適用
      const selectedColor = excelHeaderColors.find(c => c.id === excelHeaderColorValue) || excelHeaderColors[0];
      const totalRowNumber = excelShowTotal && numericColumnIndices.length > 0 
        ? filteredData.rows.length + 2 
        : filteredData.rows.length + 1;
      worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          // ヘッダー行（1行目）に色を付ける
          if (rowNumber === 1 && excelHeaderColor) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: selectedColor.argb }
            };
            cell.font = {
              bold: true,
              color: { argb: 'FF000000' }
            };
          }

          // 罫線を引く
          if (excelBorders) {
            cell.border = {
              top: { style: 'thin', color: { argb: 'FF000000' } },
              bottom: { style: 'thin', color: { argb: 'FF000000' } },
              left: { style: 'thin', color: { argb: 'FF000000' } },
              right: { style: 'thin', color: { argb: 'FF000000' } }
            };
          }

          // 数値列の数値形式を設定（小数点以下がある場合は表示）
          if (numericColumnIndices.includes(colNumber - 1)) {
            const numericIdx = numericColumnIndices.indexOf(colNumber - 1);
            const hasDecimal = hasDecimalPlaces[numericIdx];
            
            if (rowNumber > 1 && rowNumber <= totalRowNumber) {
              // データ行と合計行
              cell.numFmt = hasDecimal ? '#,##0.00' : '#,##0';
            }
          }
        });
      });

      // バッファに書き込み
      const excelBuffer = await workbook.xlsx.writeBuffer();
      zip.file(`${safeFileName}.xlsx`, excelBuffer);
    }

    const blob = await zip.generateAsync({ type: "blob" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "split_files.zip";
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-background p-4 md:p-8 transition-colors duration-300">
      <div className="mx-auto max-w-6xl">
        <div className="mb-8 relative">
          <div className="flex items-start justify-between">
            <div>
              <h1 className="text-3xl font-bold">Custom Merge Excel Web</h1>
              <p className="text-muted-foreground mt-2">高速Excelファイル統合・分割ツール</p>
            </div>
            {/* テーマ切り替えUI */}
            <div className="relative" ref={themeMenuRef}>
              <Button
                variant="outline"
                size="icon-sm"
                onClick={() => setShowThemeMenu(!showThemeMenu)}
                className="transition-all duration-200 hover:scale-105"
                title="テーマを切り替え"
              >
                <Sliders className="h-4 w-4" />
              </Button>
              {showThemeMenu && (
                <div className="absolute right-0 top-10 z-50 w-48 rounded-md border bg-popover shadow-lg p-2 space-y-1 animate-in fade-in-0 zoom-in-95">
                  <div className="px-2 py-1.5 text-xs font-semibold text-muted-foreground border-b mb-1">
                    テーマ選択
                  </div>
                  {themes.map((theme) => (
                    <button
                      key={theme.id}
                      onClick={() => {
                        setCurrentTheme(theme.id);
                        setShowThemeMenu(false);
                      }}
                      className={`w-full text-left px-3 py-2 rounded-sm text-sm transition-all duration-200 ${
                        currentTheme === theme.id
                          ? "bg-accent text-accent-foreground font-medium"
                          : "hover:bg-accent/50 text-foreground"
                      }`}
                    >
                      <div className="font-medium">{theme.name}</div>
                      <div className="text-xs text-muted-foreground">{theme.description}</div>
                    </button>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>

        <Tabs value={activeTab} onValueChange={setActiveTab}>
          <TabsList>
            <TabsTrigger value="compare">比較</TabsTrigger>
            <TabsTrigger value="split">分割</TabsTrigger>
          </TabsList>

          <TabsContent value="compare" className="mt-6">
            <Card>
              <CardHeader>
                <CardTitle>Excelファイル比較</CardTitle>
                <CardDescription>
                  2つのExcelファイルを比較し、差分を検出します
                </CardDescription>
              </CardHeader>
              <CardContent className="space-y-6">
                <div className="grid gap-4 md:grid-cols-2">
                  <div className="space-y-2">
                    <label className="text-sm font-medium">左側のファイル</label>
                    <div className="flex items-center gap-2">
                      <input
                        ref={leftFileInputRef}
                        type="file"
                        accept=".xlsx,.xls"
                        onChange={handleLeftFileChange}
                        className="hidden"
                      />
                      <Button
                        variant="outline"
                        onClick={() => leftFileInputRef.current?.click()}
                      >
                        <Upload className="mr-2 h-4 w-4" />
                        ファイルを選択
                      </Button>
                      {leftFile && (
                        <div className="flex items-center gap-2">
                          <FileSpreadsheet className="h-4 w-4" />
                          <span className="text-sm">{leftFile.name}</span>
                          <Button
                            variant="ghost"
                            size="icon-sm"
                            onClick={() => {
                              setLeftFile(null);
                              setLeftData(null);
                            }}
                          >
                            <X className="h-4 w-4" />
                          </Button>
                        </div>
                      )}
                    </div>
                    {leftData && (
                      <p className="text-xs text-muted-foreground">
                        {leftData.headers.length}列, {leftData.rows.length}行
                      </p>
                    )}
                  </div>

                  <div className="space-y-2">
                    <label className="text-sm font-medium">右側のファイル</label>
                    <div className="flex items-center gap-2">
                      <input
                        ref={rightFileInputRef}
                        type="file"
                        accept=".xlsx,.xls"
                        onChange={handleRightFileChange}
                        className="hidden"
                      />
                      <Button
                        variant="outline"
                        onClick={() => rightFileInputRef.current?.click()}
                      >
                        <Upload className="mr-2 h-4 w-4" />
                        ファイルを選択
                      </Button>
                      {rightFile && (
                        <div className="flex items-center gap-2">
                          <FileSpreadsheet className="h-4 w-4" />
                          <span className="text-sm">{rightFile.name}</span>
                          <Button
                            variant="ghost"
                            size="icon-sm"
                            onClick={() => {
                              setRightFile(null);
                              setRightData(null);
                            }}
                          >
                            <X className="h-4 w-4" />
                          </Button>
                        </div>
                      )}
                    </div>
                    {rightData && (
                      <p className="text-xs text-muted-foreground">
                        {rightData.headers.length}列, {rightData.rows.length}行
                      </p>
                    )}
                  </div>
                </div>

                {leftData && (
                  <div className="space-y-3">
                    <div className="space-y-2">
                      <label className="text-sm font-medium">①キー列（複数選択可）</label>
                      <p className="text-xs text-muted-foreground">
                        複数選択時は、選択順序が重要です。上下矢印で順序を変更できます。
                      </p>
                      <div className="max-h-[calc(100vh-500px)] min-h-[200px] overflow-y-auto rounded-md border border-input bg-background p-3 space-y-2">
                        {leftData.headers.map((header, idx) => (
                          <div key={idx} className="flex items-center space-x-2">
                            <Checkbox
                              id={`key-${idx}`}
                              checked={compareKeys.includes(header)}
                              onCheckedChange={(checked) => {
                                if (checked) {
                                  setCompareKeys([...compareKeys, header]);
                                } else {
                                  setCompareKeys(compareKeys.filter(k => k !== header));
                                }
                              }}
                            />
                            <label htmlFor={`key-${idx}`} className="text-sm font-medium leading-none cursor-pointer">
                              {header}
                            </label>
                          </div>
                        ))}
                      </div>
                    </div>
                    {compareKeys.length > 0 && (
                      <div className="space-y-2">
                        <label className="text-sm font-medium">キー列の順序（上から順に適用）</label>
                        <div className="space-y-2 rounded-md border border-input bg-background p-3">
                          {compareKeys.map((key, idx) => (
                            <div key={idx} className="flex items-center justify-between p-2 rounded-md bg-muted/50">
                              <div className="flex items-center space-x-2">
                                <span className="text-xs text-muted-foreground w-6">{idx + 1}.</span>
                                <span className="text-sm font-medium">{key}</span>
                              </div>
                              <div className="flex items-center space-x-1">
                                <Button
                                  variant="ghost"
                                  size="icon-sm"
                                  onClick={() => {
                                    if (idx > 0) {
                                      const newKeys = [...compareKeys];
                                      [newKeys[idx - 1], newKeys[idx]] = [newKeys[idx], newKeys[idx - 1]];
                                      setCompareKeys(newKeys);
                                    }
                                  }}
                                  disabled={idx === 0}
                                  className="h-6 w-6"
                                >
                                  <ChevronUp className="h-3 w-3" />
                                </Button>
                                <Button
                                  variant="ghost"
                                  size="icon-sm"
                                  onClick={() => {
                                    if (idx < compareKeys.length - 1) {
                                      const newKeys = [...compareKeys];
                                      [newKeys[idx], newKeys[idx + 1]] = [newKeys[idx + 1], newKeys[idx]];
                                      setCompareKeys(newKeys);
                                    }
                                  }}
                                  disabled={idx === compareKeys.length - 1}
                                  className="h-6 w-6"
                                >
                                  <ChevronDown className="h-3 w-3" />
                                </Button>
                                <Button
                                  variant="ghost"
                                  size="icon-sm"
                                  onClick={() => {
                                    setCompareKeys(compareKeys.filter((_, i) => i !== idx));
                                  }}
                                  className="h-6 w-6 text-destructive hover:text-destructive"
                                >
                                  <X className="h-3 w-3" />
                                </Button>
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                )}

                {leftData && rightData && compareKeys.length > 0 && (
                  <div className="space-y-2">
                    <label className="text-sm font-medium">③比較列を選択（差額計算用）</label>
                    <div className="space-y-3">
                      {compareColumns.map((col, idx) => (
                        <div key={idx} className="flex items-center gap-2 p-2 rounded-md border bg-background">
                          <div className="flex-1">
                            <select
                              value={col.left}
                              onChange={(e) => {
                                const newCols = [...compareColumns];
                                newCols[idx].left = e.target.value;
                                setCompareColumns(newCols);
                              }}
                              className="w-full rounded-md border border-input bg-background px-2 py-1 text-sm"
                            >
                              <option value="">左側の列を選択</option>
                              {leftData.headers.filter(h => !compareKeys.includes(h)).map((header, i) => (
                                <option key={i} value={header}>
                                  {header}
                                </option>
                              ))}
                            </select>
                          </div>
                          <span className="text-sm">-</span>
                          <div className="flex-1">
                            <select
                              value={col.right}
                              onChange={(e) => {
                                const newCols = [...compareColumns];
                                newCols[idx].right = e.target.value;
                                setCompareColumns(newCols);
                              }}
                              className="w-full rounded-md border border-input bg-background px-2 py-1 text-sm"
                            >
                              <option value="">右側の列を選択</option>
                              {rightData.headers.filter(h => !compareKeys.includes(h)).map((header, i) => (
                                <option key={i} value={header}>
                                  {header}
                                </option>
                              ))}
                            </select>
                          </div>
                          <input
                            type="text"
                            value={col.label}
                            onChange={(e) => {
                              const newCols = [...compareColumns];
                              newCols[idx].label = e.target.value;
                              setCompareColumns(newCols);
                            }}
                            placeholder="差額列名（例: 差額）"
                            className="w-32 rounded-md border border-input bg-background px-2 py-1 text-sm"
                          />
                          <Button
                            variant="ghost"
                            size="icon-sm"
                            onClick={() => {
                              setCompareColumns(compareColumns.filter((_, i) => i !== idx));
                            }}
                          >
                            <X className="h-4 w-4" />
                          </Button>
                        </div>
                      ))}
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => {
                          setCompareColumns([...compareColumns, { left: "", right: "", label: "差額" }]);
                        }}
                      >
                        比較列を追加
                      </Button>
                    </div>
                  </div>
                )}

                <div className="space-y-3">
                  <label className="text-sm font-medium">オプション</label>
                  <div className="flex items-center space-x-2">
                    <Checkbox
                      id="trim"
                      checked={compareOptions.trim}
                      onCheckedChange={(checked) =>
                        setCompareOptions({ ...compareOptions, trim: checked === true })
                      }
                    />
                    <label htmlFor="trim" className="text-sm font-medium leading-none">
                      前後の空白をトリム
                    </label>
                  </div>
                  <div className="flex items-center space-x-2">
                    <Checkbox
                      id="case-insensitive"
                      checked={compareOptions.case_insensitive}
                      onCheckedChange={(checked) =>
                        setCompareOptions({ ...compareOptions, case_insensitive: checked === true })
                      }
                    />
                    <label htmlFor="case-insensitive" className="text-sm font-medium leading-none">
                      大文字小文字を区別しない
                    </label>
                  </div>
                  <div className="flex items-center space-x-2">
                    <Checkbox
                      id="sort-by-keys"
                      checked={sortByKeys}
                      onCheckedChange={(checked) => setSortByKeys(checked === true)}
                    />
                    <label htmlFor="sort-by-keys" className="text-sm font-medium leading-none">
                      比較実行前にキー列でソート（推奨）
                    </label>
                  </div>
                  {sortByKeys && (
                    <p className="text-xs text-muted-foreground ml-6">
                      キー列がソートされていない場合、比較前に自動的にソートします。これにより比較処理が高速化され、結果が整理されます。
                    </p>
                  )}
                </div>

                <div className="space-y-2">
                  <Button 
                    onClick={handleCompare} 
                    disabled={!leftData || !rightData || compareKeys.length === 0}
                    className="w-full"
                  >
                    比較実行
                  </Button>
                  {(!leftData || !rightData || compareKeys.length === 0) && (
                    <p className="text-xs text-muted-foreground">
                      {!leftData && "⚠ 左側のファイルを選択してください。 "}
                      {!rightData && "⚠ 右側のファイルを選択してください。 "}
                      {compareKeys.length === 0 && "⚠ キー列を選択してください。"}
                    </p>
                  )}
                </div>

                {compareResult && mergedResult && (
                  <div className="space-y-4 rounded-lg border p-4">
                    <div>
                      <h3 className="font-semibold">比較結果</h3>
                    </div>
                    
                    {/* 列選択セクション */}
                    <div className="space-y-2">
                      <label className="text-sm font-medium">④出力する列を選択（結合キーは必須）</label>
                      <div className="max-h-[calc(100vh-500px)] min-h-[300px] overflow-y-auto rounded-md border border-input bg-background p-3 space-y-2">
                        {mergedResult.headers.map((header, idx) => {
                          // 結合キー列を検出（既に統合されているので、直接比較）
                          const isKeyColumn = compareKeys.includes(header);
                          const isChecked = selectedColumns.includes(header);
                          const isDisabled = isKeyColumn;
                          
                          return (
                            <div key={idx} className="flex items-center space-x-2">
                              <Checkbox
                                id={`col-${idx}`}
                                checked={isChecked || isDisabled}
                                disabled={isDisabled}
                                onCheckedChange={(checked) => {
                                  if (!isDisabled) {
                                    if (checked) {
                                      setSelectedColumns([...selectedColumns, header]);
                                    } else {
                                      setSelectedColumns(selectedColumns.filter(c => c !== header));
                                    }
                                  }
                                }}
                              />
                              <label 
                                htmlFor={`col-${idx}`} 
                                className={`text-sm font-medium leading-none cursor-pointer ${isDisabled ? 'text-muted-foreground' : ''}`}
                              >
                                {header}
                                {isKeyColumn && <span className="text-xs text-muted-foreground ml-1">（必須）</span>}
                              </label>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                    
                    {/* マージ結果とソートセクション */}
                    <div className="space-y-4">
                      {/* ソートセクション */}
                      <div className="space-y-2 p-3 rounded-md border bg-muted/50">
                        <label className="text-sm font-medium">ソート（最大3列、選択列のみ）</label>
                        <div className="space-y-2">
                          {[0, 1, 2].map((idx) => {
                            const sortCol = sortColumns[idx] || { column: "", direction: "asc" as const };
                            return (
                              <div key={idx} className="flex items-center gap-2">
                                <span className="text-xs text-muted-foreground w-8">{idx + 1}位:</span>
                                <select
                                  value={sortCol.column}
                                  onChange={(e) => {
                                    const newSortCols = [...sortColumns];
                                    if (e.target.value) {
                                      newSortCols[idx] = { column: e.target.value, direction: sortCol.direction };
                                      setSortColumns(newSortCols.slice(0, 3));
                                    } else {
                                      newSortCols.splice(idx, 1);
                                      setSortColumns(newSortCols);
                                    }
                                  }}
                                  className="flex-1 rounded-md border border-input bg-background px-3 py-2 text-sm"
                                >
                                  <option value="">ソート列を選択</option>
                                  {(sortedMergedResult || filterColumns(mergedResult, selectedColumns)).headers.map((header, hIdx) => (
                                    <option key={hIdx} value={header}>
                                      {header}
                                    </option>
                                  ))}
                                </select>
                                <Button
                                  variant={sortCol.direction === "desc" ? "default" : "outline"}
                                  size="sm"
                                  onClick={() => {
                                    const newSortCols = [...sortColumns];
                                    newSortCols[idx] = { ...sortCol, direction: "desc" };
                                    setSortColumns(newSortCols);
                                  }}
                                  disabled={!sortCol.column}
                                >
                                  降順
                                </Button>
                                <Button
                                  variant={sortCol.direction === "asc" ? "default" : "outline"}
                                  size="sm"
                                  onClick={() => {
                                    const newSortCols = [...sortColumns];
                                    newSortCols[idx] = { ...sortCol, direction: "asc" };
                                    setSortColumns(newSortCols);
                                  }}
                                  disabled={!sortCol.column}
                                >
                                  昇順
                                </Button>
                                {sortCol.column && (
                                  <Button
                                    variant="ghost"
                                    size="sm"
                                    onClick={() => {
                                      const newSortCols = [...sortColumns];
                                      newSortCols.splice(idx, 1);
                                      setSortColumns(newSortCols);
                                    }}
                                  >
                                    <X className="h-4 w-4" />
                                  </Button>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      </div>
                      <PreviewTable data={sortedMergedResult || filterColumns(mergedResult, selectedColumns)} />
                    </div>
                    
                    {/* Excel出力オプション */}
                    <div className="space-y-3 p-3 rounded-md border bg-muted/50">
                      <label className="text-sm font-medium">Excel出力オプション</label>
                      <div className="flex items-center space-x-2">
                        <Checkbox
                          id="excel-header-color"
                          checked={excelHeaderColor}
                          onCheckedChange={(checked) => setExcelHeaderColor(checked === true)}
                        />
                        <label htmlFor="excel-header-color" className="text-sm font-medium leading-none">
                          1行目（ヘッダー）に色を付ける
                        </label>
                      </div>
                      {excelHeaderColor && (
                        <div className="space-y-2">
                          <label className="text-xs text-muted-foreground">ヘッダー行の色</label>
                          <div className="flex flex-wrap gap-2">
                            {excelHeaderColors.map((colorOption) => (
                              <button
                                key={colorOption.id}
                                onClick={() => setExcelHeaderColorValue(colorOption.id)}
                                className={`w-10 h-10 rounded-md border-2 transition-all ${
                                  excelHeaderColorValue === colorOption.id
                                    ? "border-orange-500 scale-110 shadow-md"
                                    : "border-gray-300 hover:border-gray-400"
                                }`}
                                style={{ backgroundColor: colorOption.color }}
                                title={colorOption.name}
                              />
                            ))}
                          </div>
                          <p className="text-xs text-muted-foreground">
                            選択中: {excelHeaderColors.find(c => c.id === excelHeaderColorValue)?.name || "アクア"}
                          </p>
                        </div>
                      )}
                      <div className="flex items-center space-x-2">
                        <Checkbox
                          id="excel-borders"
                          checked={excelBorders}
                          onCheckedChange={(checked) => setExcelBorders(checked === true)}
                        />
                        <label htmlFor="excel-borders" className="text-sm font-medium leading-none">
                          罫線を引く
                        </label>
                      </div>
                      <div className="flex items-center space-x-2">
                        <Checkbox
                          id="excel-show-total"
                          checked={excelShowTotal}
                          onCheckedChange={(checked) => setExcelShowTotal(checked === true)}
                        />
                        <label htmlFor="excel-show-total" className="text-sm font-medium leading-none">
                          合計行を表示（桁区切り適用）
                        </label>
                      </div>
                    </div>
                    
                    {/* ダウンロードボタン */}
                    <div className="flex justify-end">
                      <Button variant="outline" size="sm" onClick={handleDownloadCompare}>
                        <Download className="mr-2 h-4 w-4" />
                        ダウンロード
                      </Button>
                    </div>
        </div>
                )}
              </CardContent>
            </Card>
          </TabsContent>

          <TabsContent value="split" className="mt-6">
            <Card>
              <CardHeader>
                <CardTitle>Excelファイル分割</CardTitle>
                <CardDescription>
                  キー列の値でExcelファイルを分割します
                </CardDescription>
              </CardHeader>
              <CardContent className="space-y-6">
                <div className="space-y-2">
                  <label className="text-sm font-medium">ファイル</label>
                  <div className="flex items-center gap-2">
                    <input
                      ref={splitFileInputRef}
                      type="file"
                      accept=".xlsx,.xls"
                      onChange={handleSplitFileChange}
                      className="hidden"
                    />
                    <Button
                      variant="outline"
                      onClick={() => splitFileInputRef.current?.click()}
                    >
                      <Upload className="mr-2 h-4 w-4" />
                      ファイルを選択
                    </Button>
                    {splitFile && (
                      <div className="flex items-center gap-2">
                        <FileSpreadsheet className="h-4 w-4" />
                        <span className="text-sm">{splitFile.name}</span>
                        <Button
                          variant="ghost"
                          size="icon-sm"
                          onClick={() => {
                            setSplitFile(null);
                            setSplitData(null);
                            setSplitKeys([]);
                            setSplitResult(null);
                          }}
                        >
                          <X className="h-4 w-4" />
                        </Button>
                      </div>
                    )}
                  </div>
                  {splitData && (
                    <p className="text-xs text-muted-foreground">
                      {splitData.headers.length}列, {splitData.rows.length}行
                    </p>
                  )}
                </div>

                {splitData && (
                  <div className="space-y-2">
                    <label className="text-sm font-medium">キー列（複数選択可）</label>
                    <div className="max-h-[calc(100vh-400px)] min-h-[300px] overflow-y-auto rounded-md border border-input bg-background p-3 space-y-2">
                      {splitData.headers.map((header, idx) => (
                        <div key={idx} className="flex items-center space-x-2">
                          <Checkbox
                            id={`split-key-${idx}`}
                            checked={splitKeys.includes(header)}
                            onCheckedChange={(checked) => {
                              if (checked) {
                                setSplitKeys([...splitKeys, header]);
                              } else {
                                setSplitKeys(splitKeys.filter(k => k !== header));
                              }
                            }}
                          />
                          <label htmlFor={`split-key-${idx}`} className="text-sm font-medium leading-none cursor-pointer">
                            {header}
                          </label>
                        </div>
                      ))}
                    </div>
                    {splitKeys.length > 0 && (
                      <p className="text-xs text-muted-foreground">
                        選択中: {splitKeys.join(", ")}
                      </p>
                    )}
                  </div>
                )}

                <Button onClick={handleSplit} disabled={!splitData || splitKeys.length === 0}>
                  分割実行
                </Button>

                {splitResult && splitResult.parts.length > 0 && (
                  <div className="space-y-4 rounded-lg border p-4">
                    <div>
                      <h3 className="font-semibold">分割結果</h3>
                    </div>
                    <div className="text-sm">
                      {splitResult.parts.length}個のファイルに分割されました
                    </div>
                    
                    {/* 列選択セクション */}
                    <div className="space-y-2">
                      <label className="text-sm font-medium">出力する列を選択（結合キーは必須）</label>
                      <div className="max-h-[calc(100vh-500px)] min-h-[300px] overflow-y-auto rounded-md border border-input bg-background p-3 space-y-2">
                        {splitResult.parts[0].table.headers.map((header: string, idx: number) => {
                          const isKeyColumn = splitKeys.includes(header);
                          const isChecked = selectedSplitColumns.includes(header);
                          const isNumeric = splitNumericColumns.includes(header);
                          const isDisabled = isKeyColumn;
                          
                          return (
                            <div key={idx} className="flex items-center space-x-2">
                              <Checkbox
                                id={`split-col-${idx}`}
                                checked={isChecked || isDisabled}
                                disabled={isDisabled}
                                onCheckedChange={(checked) => {
                                  if (!isDisabled) {
                                    if (checked) {
                                      setSelectedSplitColumns([...selectedSplitColumns, header]);
                                    } else {
                                      setSelectedSplitColumns(selectedSplitColumns.filter(c => c !== header));
                                      // 数値列の選択も解除
                                      setSplitNumericColumns(splitNumericColumns.filter(c => c !== header));
                                    }
                                  }
                                }}
                              />
                              {isChecked && !isKeyColumn && (
                                <Checkbox
                                  id={`split-numeric-${idx}`}
                                  checked={isNumeric}
                                  onCheckedChange={(checked) => {
                                    if (checked) {
                                      setSplitNumericColumns([...splitNumericColumns, header]);
                                    } else {
                                      setSplitNumericColumns(splitNumericColumns.filter(c => c !== header));
                                    }
                                  }}
                                  className="ml-2"
                                />
                              )}
                              <label 
                                htmlFor={`split-col-${idx}`} 
                                className={`text-sm font-medium leading-none cursor-pointer ${isDisabled ? 'text-muted-foreground' : ''}`}
                              >
                                {header}
                                {isKeyColumn && <span className="text-xs text-muted-foreground ml-1">（必須）</span>}
                                {isChecked && !isKeyColumn && (
                                  <span className="text-xs text-muted-foreground ml-2">
                                    <label htmlFor={`split-numeric-${idx}`} className="cursor-pointer">
                                      （数値列: {isNumeric ? '✓' : 'なし'}）
                                    </label>
                                  </span>
                                )}
                              </label>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                    
                    {/* プレビューセクション */}
                    <div className="space-y-2">
                      <label className="text-sm font-medium">プレビュー</label>
                      <div className="max-h-96 overflow-auto rounded-md border">
                        <table className="w-full text-sm">
                          <thead className="bg-muted sticky top-0">
                            <tr>
                              <th className="px-4 py-2 text-left font-medium border-b">ファイル名</th>
                              <th className="px-4 py-2 text-left font-medium border-b">行数</th>
                            </tr>
                          </thead>
                          <tbody>
                            {splitResult.parts.slice(0, 100).map((part: any, idx: number) => (
                              <tr key={idx} className="border-b hover:bg-muted/50">
                                <td className="px-4 py-2 max-w-xs truncate" title={part.key_value}>
                                  {part.key_value}
                                </td>
                                <td className="px-4 py-2">{part.table.rows.length}行</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                        {splitResult.parts.length > 100 && (
                          <div className="p-2 text-xs text-muted-foreground text-center border-t">
                            最初の100件を表示しています（全{splitResult.parts.length}件）
                          </div>
                        )}
                      </div>
                    </div>
                    
                    {/* データプレビュー（最初の5個のファイルの内容を表示） */}
                    {splitResult.parts.length > 0 && (
                      <div className="space-y-4">
                        <label className="text-sm font-medium">データプレビュー（最初の5個のファイル）</label>
                        
                        {splitResult.parts.slice(0, 5).map((part: any, partIdx: number) => {
                          // 選択された列でフィルタリング
                          const filteredData = filterColumns(part.table, selectedSplitColumns);
                          
                          // ソート処理
                          let sortedData = filteredData;
                          if (splitSortColumns.length > 0) {
                            const sortedRows = [...filteredData.rows].sort((a, b) => {
                              for (const sortCol of splitSortColumns) {
                                const sortIdx = filteredData.headers.indexOf(sortCol.column);
                                if (sortIdx === -1) continue;
                                
                                const aVal = parseFloat(a[sortIdx]) || 0;
                                const bVal = parseFloat(b[sortIdx]) || 0;
                                if (isNaN(aVal) || isNaN(bVal)) {
                                  const aStr = String(a[sortIdx] || "");
                                  const bStr = String(b[sortIdx] || "");
                                  const result = sortCol.direction === "asc" ? aStr.localeCompare(bStr) : bStr.localeCompare(aStr);
                                  if (result !== 0) return result;
                                } else {
                                  const result = sortCol.direction === "asc" ? aVal - bVal : bVal - aVal;
                                  if (result !== 0) return result;
                                }
                              }
                              return 0;
                            });
                            sortedData = {
                              ...filteredData,
                              rows: sortedRows,
                            };
                          }
                          
                          return (
                            <div key={partIdx} className="space-y-2">
                              <div className="text-sm font-medium border-b pb-1">
                                {part.key_value} ({part.table.rows.length}行)
                              </div>
                              <PreviewTable data={sortedData} />
                            </div>
                          );
                        })}
                        
                        {/* ソートセクション（全ファイル共通） */}
                        <div className="space-y-2 p-3 rounded-md border bg-muted/50">
                          <label className="text-sm font-medium">ソート（最大3列、選択列のみ）</label>
                          <div className="space-y-2">
                            {[0, 1, 2].map((idx) => {
                              const sortCol = splitSortColumns[idx] || { column: "", direction: "asc" as const };
                              return (
                                <div key={idx} className="flex items-center gap-2">
                                  <span className="text-xs text-muted-foreground w-8">{idx + 1}位:</span>
                                  <select
                                    value={sortCol.column}
                                    onChange={(e) => {
                                      const newSortCols = [...splitSortColumns];
                                      if (e.target.value) {
                                        newSortCols[idx] = { column: e.target.value, direction: sortCol.direction };
                                        setSplitSortColumns(newSortCols.slice(0, 3));
                                      } else {
                                        newSortCols.splice(idx, 1);
                                        setSplitSortColumns(newSortCols);
                                      }
                                    }}
                                    className="flex-1 rounded-md border border-input bg-background px-3 py-2 text-sm"
                                  >
                                    <option value="">ソート列を選択</option>
                                    {splitResult.parts.length > 0 && filterColumns(splitResult.parts[0].table, selectedSplitColumns).headers.map((header: string, hIdx: number) => (
                                      <option key={hIdx} value={header}>
                                        {header}
                                      </option>
                                    ))}
                                  </select>
                                  <Button
                                    variant={sortCol.direction === "desc" ? "default" : "outline"}
                                    size="sm"
                                    onClick={() => {
                                      const newSortCols = [...splitSortColumns];
                                      newSortCols[idx] = { ...sortCol, direction: "desc" };
                                      setSplitSortColumns(newSortCols);
                                    }}
                                    disabled={!sortCol.column}
                                  >
                                    降順
                                  </Button>
                                  <Button
                                    variant={sortCol.direction === "asc" ? "default" : "outline"}
                                    size="sm"
                                    onClick={() => {
                                      const newSortCols = [...splitSortColumns];
                                      newSortCols[idx] = { ...sortCol, direction: "asc" };
                                      setSplitSortColumns(newSortCols);
                                    }}
                                    disabled={!sortCol.column}
                                  >
                                    昇順
                                  </Button>
                                  {sortCol.column && (
                                    <Button
                                      variant="ghost"
                                      size="sm"
                                      onClick={() => {
                                        const newSortCols = [...splitSortColumns];
                                        newSortCols.splice(idx, 1);
                                        setSplitSortColumns(newSortCols);
                                      }}
                                    >
                                      <X className="h-4 w-4" />
                                    </Button>
                                  )}
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      </div>
                    )}
                    
                    {/* Excel出力オプション */}
                    <div className="space-y-3 p-3 rounded-md border bg-muted/50">
                      <label className="text-sm font-medium">Excel出力オプション</label>
                      <div className="flex items-center space-x-2">
                        <Checkbox
                          id="excel-header-color-split"
                          checked={excelHeaderColor}
                          onCheckedChange={(checked) => setExcelHeaderColor(checked === true)}
                        />
                        <label htmlFor="excel-header-color-split" className="text-sm font-medium leading-none">
                          1行目（ヘッダー）に色を付ける
                        </label>
                      </div>
                      {excelHeaderColor && (
                        <div className="space-y-2">
                          <label className="text-xs text-muted-foreground">ヘッダー行の色</label>
                          <div className="flex flex-wrap gap-2">
                            {excelHeaderColors.map((colorOption) => (
                              <button
                                key={colorOption.id}
                                onClick={() => setExcelHeaderColorValue(colorOption.id)}
                                className={`w-10 h-10 rounded-md border-2 transition-all ${
                                  excelHeaderColorValue === colorOption.id
                                    ? "border-orange-500 scale-110 shadow-md"
                                    : "border-gray-300 hover:border-gray-400"
                                }`}
                                style={{ backgroundColor: colorOption.color }}
                                title={colorOption.name}
                              />
                            ))}
                          </div>
                          <p className="text-xs text-muted-foreground">
                            選択中: {excelHeaderColors.find(c => c.id === excelHeaderColorValue)?.name || "アクア"}
                          </p>
                        </div>
                      )}
                      <div className="flex items-center space-x-2">
                        <Checkbox
                          id="excel-borders-split"
                          checked={excelBorders}
                          onCheckedChange={(checked) => setExcelBorders(checked === true)}
                        />
                        <label htmlFor="excel-borders-split" className="text-sm font-medium leading-none">
                          罫線を引く
                        </label>
                      </div>
                      <div className="flex items-center space-x-2">
                        <Checkbox
                          id="excel-show-total-split"
                          checked={excelShowTotal}
                          onCheckedChange={(checked) => setExcelShowTotal(checked === true)}
                        />
                        <label htmlFor="excel-show-total-split" className="text-sm font-medium leading-none">
                          合計行を表示（桁区切り適用）
                        </label>
                      </div>
                    </div>
                    
                    {/* ダウンロードボタン */}
                    <div className="flex justify-end">
                      <Button variant="outline" size="sm" onClick={handleDownloadSplit}>
                        <Download className="mr-2 h-4 w-4" />
                        ZIPでダウンロード
                      </Button>
                    </div>
                  </div>
                )}
              </CardContent>
            </Card>
          </TabsContent>
        </Tabs>
        </div>
    </div>
  );
}
