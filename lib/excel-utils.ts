import * as XLSX from 'xlsx';

export interface TableData {
  headers: string[];
  rows: string[][];
}

export function readExcelFile(file: File): Promise<TableData> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' }) as any[][];

        if (jsonData.length === 0) {
          reject(new Error('ファイルが空です'));
          return;
        }

        const headers = jsonData[0].map((cell: any) => String(cell || ''));
        const rows = jsonData.slice(1).map((row: any[]) => row.map((cell: any) => String(cell || '')));

        resolve({ headers, rows });
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

export function writeExcelFile(data: TableData, filename: string): void {
  const worksheet = XLSX.utils.aoa_to_sheet([data.headers, ...data.rows]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  XLSX.writeFile(workbook, filename);
}

