import { describe, expect, test } from 'bun:test';
import fs from 'node:fs';
import path from 'node:path';
import * as XLSX from 'xlsx';
import { filterDbRows, parseCrmWorkbook } from '../lib/excelParser';
import {
  buildDefaultExportFileName,
  buildExportWorkbookBuffer,
  resolveExportDbRows,
} from '../lib/excelExporter';

describe('excelExporter', () => {
  const testXlsxPath = path.resolve(process.cwd(), 'test.xlsx');
  const fileBuffer = fs.readFileSync(testXlsxPath);
  const workbook = parseCrmWorkbook(fileBuffer, 'test.xlsx');

  test('retains only ancestor section headers in original row order for target records', () => {
    // Pick row 16 (MA001-02) and row 50 (MZ001), intentionally passed out of order
    const summary = resolveExportDbRows(workbook, {
      targetRecordRowNumbers: new Set([50, 16]),
    });

    expect(summary.recordCount).toBe(2);
    expect(summary.sectionCount).toBe(5); // MA000 (2), MA001 (3), MA001-02 (15), MZ000 (48), MZ001 (49)
    expect(summary.rowsToExport.map((r) => r.rowNumber)).toEqual([2, 3, 15, 16, 48, 49, 50]);
  });

  test('preserves original row order in export even when source records were sorted in UI', () => {
    const sortedRecords = filterDbRows(workbook.dbRows, {
      searchQuery: '',
      attendanceFilter: 'all',
      groupFilter: 'MZ001',
      showSectionHeaders: true,
      sort: { column: '出席活動', direction: 'desc' },
    });
    // UI sorted order has row 51 before row 50
    expect(sortedRecords[0]?.rowNumber).toBe(51);

    const summary = resolveExportDbRows(workbook, {
      targetRecordRowNumbers: new Set(sortedRecords.map((r) => r.rowNumber)),
    });
    // Export order strictly follows original rowNumber ascending: 48 (MZ000), 49 (MZ001), 50..59
    expect(summary.rowsToExport.map((r) => r.rowNumber)).toEqual([
      48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59,
    ]);
  });

  test('builds valid .xlsx buffer preserving all 50 DB headers, G:H section merges, 組別 and 界別 sheets', () => {
    const outBuffer = buildExportWorkbookBuffer(workbook, {
      targetRecordRowNumbers: new Set([5, 16, 50]),
    });

    const rawExported = XLSX.read(outBuffer, { type: 'array' });
    expect(rawExported.SheetNames).toEqual(['DB', '組別', '界別']);
    const dbSheet = rawExported.Sheets['DB'];
    expect(dbSheet['!merges']?.length).toBe(6); // 6 section rows exported: 2, 3, 4, 15, 48, 49
    expect(dbSheet['!merges']?.[0]).toEqual({
      s: { c: 6, r: 1 },
      e: { c: 7, r: 1 },
    });

    // Re-import round-trip verification
    const reimported = parseCrmWorkbook(outBuffer, 'exported.xlsx');
    expect(reimported.dbHeaders).toEqual(workbook.dbHeaders);
    expect(reimported.summary.totalRecords).toBe(3);
    expect(reimported.summary.sectionCount).toBe(6);
    expect(reimported.groups.length).toBe(558);
    expect(reimported.sectors.length).toBe(26);
  });

  test('formats default export file name cleanly', () => {
    const fixedDate = new Date(2026, 9, 31, 14, 5);
    expect(buildDefaultExportFileName('test.xlsx', '篩選匯出', fixedDate)).toBe(
      'test_篩選匯出_20261031_1405.xlsx',
    );
  });

  test('exports custom subset of DB columns plus extra event columns (blank by default, or populated when enabled)', () => {
    const customHeaders = ['組別', '姓名', '職位', '桌號', '車牌'];
    const extraMap = new Map<number, Record<string, string>>([
      [5, { 桌號: 'VIP-1', 車牌: 'MA-00-01' }],
      [50, { 桌號: 'B-12', 車牌: 'MZ-99-99' }],
    ]);

    // Case 1: populateExtraColumnValues = false (default: blank extra columns)
    const blankExtraBuf = buildExportWorkbookBuffer(workbook, {
      targetRecordRowNumbers: new Set([5, 50]),
      outputHeaders: customHeaders,
      populateExtraColumnValues: false,
      extraValuesByRowNumber: extraMap,
    });
    const wbBlank = XLSX.read(blankExtraBuf, { type: 'array' });
    const rowsBlank = XLSX.utils.sheet_to_json<string[]>(wbBlank.Sheets['DB'], {
      header: 1,
      defval: '',
    });
    expect(rowsBlank[0]).toEqual(customHeaders);
    // Section merge shifts to columns 1..2 (姓名..職位)
    expect(wbBlank.Sheets['DB']['!merges']?.[0]).toEqual({
      s: { c: 1, r: 1 },
      e: { c: 2, r: 1 },
    });
    // Record row 5 is at index 4 (after header + 3 ancestor sections 2,3,4)
    expect(rowsBlank[4]?.[3]).toBe('');
    expect(rowsBlank[4]?.[4]).toBe('');

    // Case 2: populateExtraColumnValues = true (carries over extra values for matched records)
    const filledExtraBuf = buildExportWorkbookBuffer(workbook, {
      targetRecordRowNumbers: new Set([5, 50]),
      outputHeaders: customHeaders,
      populateExtraColumnValues: true,
      extraValuesByRowNumber: extraMap,
    });
    const wbFilled = XLSX.read(filledExtraBuf, { type: 'array' });
    const rowsFilled = XLSX.utils.sheet_to_json<string[]>(wbFilled.Sheets['DB'], {
      header: 1,
      defval: '',
    });
    expect(rowsFilled[4]?.[3]).toBe('VIP-1');
    expect(rowsFilled[4]?.[4]).toBe('MA-00-01');
  });

  test('falls back section header label to first column when 姓名 is omitted from outputHeaders (even when Remark is non-empty)', () => {
    const noNameHeaders = ['組別', 'PID', '出席活動', 'Remark'];
    const buf = buildExportWorkbookBuffer(workbook, {
      targetRecordRowNumbers: new Set([50]),
      outputHeaders: noNameHeaders,
    });
    const wbOut = XLSX.read(buf, { type: 'array' });
    const matrix = XLSX.utils.sheet_to_json<string[]>(wbOut.Sheets['DB'], {
      header: 1,
      defval: '',
    });
    expect(matrix[0]).toEqual(noNameHeaders);
    expect(String(matrix[1]?.[0] ?? '')).toContain('MZ000');
    // Row 49 (MZ001 基本法委員會) has a non-empty Remark in test.xlsx; column 0 must still get the fallback section label
    expect(String(matrix[2]?.[0] ?? '')).toContain('MZ001');
  });
});
