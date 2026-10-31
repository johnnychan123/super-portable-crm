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

  test('applies test.xlsx column widths, table styles, section fills, red non-attending rows, and numeric types across DB, 組別, and 界別', () => {
    // Export records 5 (MA001-01 attending, 計算=1, under MA000 -> MA001 -> MA001-01),
    // 28 (MA002 not-attending, 計算=1, under MA000 -> MA002), and 58 (MZ001 not-attending, 計算=0, under MZ000 -> MZ001)
    const outBuffer = buildExportWorkbookBuffer(workbook, {
      targetRecordRowNumbers: new Set([5, 28, 58]),
    });

    // 1. Verify column widths via SheetJS cellStyles/cols reader
    const exportedWb = XLSX.read(outBuffer, { type: 'array', cellStyles: true });
    const dbCols = exportedWb.Sheets['DB']['!cols'] ?? [];
    expect(dbCols[0]?.width).toBe(5); // A: 酒會
    expect(dbCols[1]?.width).toBe(5); // B: 升旗
    expect(dbCols[2]?.width).toBe(5); // C: 計算
    expect(dbCols[3]?.width).toBe(11); // D: 組別
    expect(dbCols[6]?.width).toBe(25); // G: 姓名
    expect(dbCols[7]?.width).toBe(60); // H: 職位

    const groupCols = exportedWb.Sheets['組別']['!cols'] ?? [];
    expect(groupCols[1]?.width).toBe(10); // B: 界別
    expect(groupCols[2]?.width).toBe(10); // C: 分類
    expect(groupCols[3]?.width).toBe(10); // D: 簡稱
    expect(groupCols[4]?.width).toBe(60); // E: 中文
    expect(groupCols[5]?.width).toBe(60); // F

    const sectorCols = exportedWb.Sheets['界別']['!cols'] ?? [];
    expect(sectorCols[1]?.width).toBe(10); // B: 簡稱
    expect(sectorCols[2]?.width).toBe(25); // C: 中文

    // 2. Verify numeric cell types for DB 計算 and 組別/界別 ID
    // Exported DB rows:
    // Row 1: Header
    // Row 2: MA000 公共行政部門 (upper section -> lightGray)
    // Row 3: MA001 XXXX (parent group section -> lightGray)
    // Row 4: MA001-01 XXXX (leaf section -> gray125)
    // Row 5: Record from row 5 (MA001-01, attending 'Y', 計算=1)
    // Row 6: MA002 XXXX (leaf section -> gray125)
    // Row 7: Record from row 28 (MA002, not-attending '不出席活動', 計算=1)
    // Row 8: MZ000 委員會 (upper section -> lightGray)
    // Row 9: MZ001 基本法委員會 (leaf section -> gray125)
    // Row 10: Record from row 58 (MZ001, not-attending '不出席活動', 計算=0)
    expect(exportedWb.Sheets['DB']['C5']?.t).toBe('n');
    expect(exportedWb.Sheets['DB']['C5']?.v).toBe(1);
    expect(exportedWb.Sheets['DB']['C7']?.t).toBe('n');
    expect(exportedWb.Sheets['DB']['C7']?.v).toBe(1);
    expect(exportedWb.Sheets['DB']['C10']?.t).toBe('n');
    expect(exportedWb.Sheets['DB']['C10']?.v).toBe(0);
    expect(exportedWb.Sheets['組別']['A2']?.t).toBe('n');
    expect(exportedWb.Sheets['組別']['A2']?.v).toBe(596);
    expect(exportedWb.Sheets['界別']['A2']?.t).toBe('n');
    expect(exportedWb.Sheets['界別']['A2']?.v).toBe(2);

    // 3. Verify OpenXML styles.xml and worksheet XML cell style indices (s="...")
    const CFB = (XLSX as unknown as { CFB: typeof XLSX.CFB }).CFB;
    const cfb = CFB.read(outBuffer, { type: 'array' });
    const decoder = new TextDecoder();
    const stylesXml = decoder.decode(
      new Uint8Array(CFB.find(cfb, '/xl/styles.xml')!.content as ArrayBuffer),
    );
    expect(stylesXml).toContain('name val="Times New Roman"');
    expect(stylesXml).toContain('name val="新細明體"');
    expect(stylesXml).toContain('patternType="lightGray"');
    expect(stylesXml).toContain('patternType="gray125"');
    expect(stylesXml).toContain('color indexed="10"');
    expect(stylesXml).toContain('defaultTableStyle="TableStyleMedium2"');

    const sheet1Xml = decoder.decode(
      new Uint8Array(CFB.find(cfb, '/xl/worksheets/sheet1.xml')!.content as ArrayBuffer),
    );
    // Header row uses s="3"
    expect(sheet1Xml).toMatch(/<c r="A1"[^>]*s="3"/);
    expect(sheet1Xml).toMatch(/<c r="AX1"[^>]*s="3"/);
    // Row 2 (MA000) & Row 3 (MA001) upper-level sections use s="4" and G uses s="12"
    expect(sheet1Xml).toMatch(/<c r="A2"[^>]*s="4"/);
    expect(sheet1Xml).toMatch(/<c r="G2"[^>]*s="12"/);
    expect(sheet1Xml).toMatch(/<c r="AX2"[^>]*s="4"/);
    expect(sheet1Xml).toMatch(/<c r="A3"[^>]*s="4"/);
    expect(sheet1Xml).toMatch(/<c r="G3"[^>]*s="12"/);
    // Row 4 (MA001-01) & Row 6 (MA002) leaf sections use s="5" and G uses s="14"
    expect(sheet1Xml).toMatch(/<c r="A4"[^>]*s="5"/);
    expect(sheet1Xml).toMatch(/<c r="G4"[^>]*s="14"/);
    expect(sheet1Xml).toMatch(/<c r="AX4"[^>]*s="5"/);
    expect(sheet1Xml).toMatch(/<c r="A6"[^>]*s="5"/);
    expect(sheet1Xml).toMatch(/<c r="G6"[^>]*s="14"/);
    // Row 5 (attending record) uses s="6" on A..C and s="1" on D..AX
    expect(sheet1Xml).toMatch(/<c r="A5"[^>]*s="6"/);
    expect(sheet1Xml).toMatch(/<c r="C5"[^>]*s="6"/);
    expect(sheet1Xml).toMatch(/<c r="D5"[^>]*s="1"/);
    expect(sheet1Xml).toMatch(/<c r="AX5"[^>]*s="1"/);
    // Row 7 (not-attending record) uses s="7" on A..C (centered red) and s="2" on D..AX (red)
    expect(sheet1Xml).toMatch(/<c r="A7"[^>]*s="7"/);
    expect(sheet1Xml).toMatch(/<c r="C7"[^>]*s="7"/);
    expect(sheet1Xml).toMatch(/<c r="D7"[^>]*s="2"/);
    expect(sheet1Xml).toMatch(/<c r="E7"[^>]*s="2"/);
    expect(sheet1Xml).toMatch(/<c r="AX7"[^>]*s="2"/);
    // Frozen top row and default row height 15.75
    expect(sheet1Xml).toContain('defaultRowHeight="15.75"');
    expect(sheet1Xml).toContain('state="frozen"');

    // Sheet 2 (組別) styles: header A1..D1 s="9", E1 s="3", data rows s="8"
    const sheet2Xml = decoder.decode(
      new Uint8Array(CFB.find(cfb, '/xl/worksheets/sheet2.xml')!.content as ArrayBuffer),
    );
    expect(sheet2Xml).toMatch(/<c r="A1"[^>]*s="9"/);
    expect(sheet2Xml).toMatch(/<c r="E1"[^>]*s="3"/);
    expect(sheet2Xml).toMatch(/<c r="A2"[^>]*s="8"/);
    expect(sheet2Xml).toMatch(/<c r="D2"[^>]*s="8"/);
    expect(sheet2Xml).toContain('defaultRowHeight="15.75"');
    expect(sheet2Xml).toContain('state="frozen"');

    // Sheet 3 (界別) styles: header A1..D1 s="11", data rows s="10"
    const sheet3Xml = decoder.decode(
      new Uint8Array(CFB.find(cfb, '/xl/worksheets/sheet3.xml')!.content as ArrayBuffer),
    );
    expect(sheet3Xml).toMatch(/<c r="A1"[^>]*s="11"/);
    expect(sheet3Xml).toMatch(/<c r="D1"[^>]*s="11"/);
    expect(sheet3Xml).toMatch(/<c r="A2"[^>]*s="10"/);
    expect(sheet3Xml).toMatch(/<c r="D2"[^>]*s="10"/);
    expect(sheet3Xml).toContain('defaultRowHeight="15.75"');
    expect(sheet3Xml).toContain('state="frozen"');
  });

  test('maps DB column widths and centered/section styles dynamically when outputHeaders order is customized', () => {
    const customHeaders = ['組別', '姓名', '職位', '酒會', '計算', '出席活動', '桌號'];
    const outBuffer = buildExportWorkbookBuffer(workbook, {
      targetRecordRowNumbers: new Set([5, 28]),
      outputHeaders: customHeaders,
    });

    const exportedWb = XLSX.read(outBuffer, { type: 'array', cellStyles: true });
    const dbCols = exportedWb.Sheets['DB']['!cols'] ?? [];
    expect(dbCols[0]?.width).toBe(11); // A: 組別
    expect(dbCols[1]?.width).toBe(25); // B: 姓名
    expect(dbCols[2]?.width).toBe(60); // C: 職位
    expect(dbCols[3]?.width).toBe(5); // D: 酒會
    expect(dbCols[4]?.width).toBe(5); // E: 計算

    const CFB = (XLSX as unknown as { CFB: typeof XLSX.CFB }).CFB;
    const cfb = CFB.read(outBuffer, { type: 'array' });
    const decoder = new TextDecoder();
    const sheet1Xml = decoder.decode(
      new Uint8Array(CFB.find(cfb, '/xl/worksheets/sheet1.xml')!.content as ArrayBuffer),
    );
    // Row 2 is MA000 (upper section): B2 (姓名) has s="12", A2 and G2 have s="4"
    expect(sheet1Xml).toMatch(/<c r="A2"[^>]*s="4"/);
    expect(sheet1Xml).toMatch(/<c r="B2"[^>]*s="12"/);
    expect(sheet1Xml).toMatch(/<c r="G2"[^>]*s="4"/);
    // Row 5 is attending record: D5 (酒會) and E5 (計算) have centered s="6", A5 (組別) has s="1"
    expect(sheet1Xml).toMatch(/<c r="A5"[^>]*s="1"/);
    expect(sheet1Xml).toMatch(/<c r="D5"[^>]*s="6"/);
    expect(sheet1Xml).toMatch(/<c r="E5"[^>]*s="6"/);
    // Row 7 is not-attending record: D7 (酒會) and E7 (計算) have centered red s="7", A7 (組別) has red s="2"
    expect(sheet1Xml).toMatch(/<c r="A7"[^>]*s="2"/);
    expect(sheet1Xml).toMatch(/<c r="D7"[^>]*s="7"/);
    expect(sheet1Xml).toMatch(/<c r="E7"[^>]*s="7"/);
  });
});
