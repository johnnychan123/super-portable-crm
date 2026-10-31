import * as XLSX from 'xlsx';
import type { DbRowItem, ParsedCrmWorkbook } from './excelParser';

export interface ExportCrmOptions {
  targetRecordRowNumbers: ReadonlySet<number>;
  outputHeaders?: string[];
  populateExtraColumnValues?: boolean;
  extraValuesByRowNumber?: ReadonlyMap<number, Record<string, string>>;
}

export interface ExportPreviewSummary {
  rowsToExport: DbRowItem[];
  recordCount: number;
  sectionCount: number;
}

export function resolveExportDbRows(
  workbook: ParsedCrmWorkbook,
  options: ExportCrmOptions,
): ExportPreviewSummary {
  const includedRowNumbers = new Set<number>();
  let recordCount = 0;

  for (const row of workbook.dbRows) {
    if (row.kind === 'record' && options.targetRecordRowNumbers.has(row.rowNumber)) {
      includedRowNumbers.add(row.rowNumber);
      recordCount++;
      for (const ancRowNum of row.ancestorSectionRows ?? []) {
        includedRowNumbers.add(ancRowNum);
      }
    }
  }

  const rowsToExport = workbook.dbRows
    .filter((row) => includedRowNumbers.has(row.rowNumber))
    .sort((a, b) => a.rowNumber - b.rowNumber);

  const sectionCount = rowsToExport.filter((r) => r.kind === 'section').length;

  return {
    rowsToExport,
    recordCount,
    sectionCount,
  };
}

export function buildExportWorkbookBuffer(
  workbook: ParsedCrmWorkbook,
  options: ExportCrmOptions,
): Uint8Array {
  const { rowsToExport } = resolveExportDbRows(workbook, options);
  const outWb = XLSX.utils.book_new();

  const headersToUse =
    options.outputHeaders && options.outputHeaders.length > 0
      ? options.outputHeaders
      : workbook.dbHeaders;
  const dbHeaderSet = new Set(workbook.dbHeaders);
  const nameColIdx = headersToUse.indexOf('姓名');

  // 1. Build DB worksheet with selected headers in order
  const dbMatrix: string[][] = [headersToUse];
  const sectionMerges: XLSX.Range[] = [];

  for (let i = 0; i < rowsToExport.length; i++) {
    const row = rowsToExport[i];
    const outRowIndex = i + 1; // 0-based index in worksheet (row 0 is header)

    if (row.kind === 'section') {
      const rowCells = headersToUse.map((header) =>
        dbHeaderSet.has(header) ? (row.values[header] ?? '') : '',
      );

      if (nameColIdx >= 0) {
        if (headersToUse[nameColIdx + 1] === '職位') {
          sectionMerges.push({
            s: { c: nameColIdx, r: outRowIndex },
            e: { c: nameColIdx + 1, r: outRowIndex },
          });
        }
      } else if (rowCells.length > 0 && rowCells.every((c) => c === '')) {
        rowCells[0] =
          row.values['姓名'] || row.currentSectionLabel || row.sectionCode || '';
      }

      dbMatrix.push(rowCells);
    } else {
      const extraRowVals = options.populateExtraColumnValues
        ? options.extraValuesByRowNumber?.get(row.rowNumber)
        : undefined;
      const rowCells = headersToUse.map((header) => {
        if (dbHeaderSet.has(header)) {
          return row.values[header] ?? '';
        }
        return extraRowVals?.[header] ?? '';
      });
      dbMatrix.push(rowCells);
    }
  }

  const dbSheet = XLSX.utils.aoa_to_sheet(dbMatrix);
  if (sectionMerges.length > 0) {
    dbSheet['!merges'] = sectionMerges;
  }
  XLSX.utils.book_append_sheet(outWb, dbSheet, 'DB');

  // 2. Build 組別 worksheet preserving original rows
  const groupMatrix: string[][] =
    workbook.rawGroupRows && workbook.rawGroupRows.length > 0
      ? workbook.rawGroupRows
      : [
          ['ID', '界別', '分類', '簡稱', '中文名稱'],
          ...workbook.groups.map((g) => [g.id, g.sectorCode, g.category, g.code, g.nameZh]),
        ];
  const groupSheet = XLSX.utils.aoa_to_sheet(groupMatrix);
  XLSX.utils.book_append_sheet(outWb, groupSheet, '組別');

  // 3. Build 界別 worksheet preserving original rows
  const sectorMatrix: string[][] =
    workbook.rawSectorRows && workbook.rawSectorRows.length > 0
      ? workbook.rawSectorRows
      : [
          ['ID', '簡稱', '中文名稱', '備註'],
          ...workbook.sectors.map((s) => [s.id, s.code, s.nameZh, s.remark]),
        ];
  const sectorSheet = XLSX.utils.aoa_to_sheet(sectorMatrix);
  XLSX.utils.book_append_sheet(outWb, sectorSheet, '界別');

  const arrayOut = XLSX.write(outWb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
  return new Uint8Array(arrayOut);
}

export function buildDefaultExportFileName(
  originalFileName: string,
  suffix = '篩選匯出',
  now: Date = new Date(),
): string {
  const baseName = originalFileName.replace(/\.(xlsx|xls)$/i, '') || 'crm_export';
  const yyyy = String(now.getFullYear());
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  const hh = String(now.getHours()).padStart(2, '0');
  const min = String(now.getMinutes()).padStart(2, '0');
  return `${baseName}_${suffix}_${yyyy}${mm}${dd}_${hh}${min}.xlsx`;
}
