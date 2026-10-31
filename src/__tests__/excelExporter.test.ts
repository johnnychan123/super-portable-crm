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
});
