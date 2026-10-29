import { describe, expect, test } from 'bun:test';
import fs from 'node:fs';
import path from 'node:path';
import {
  filterDbRows,
  filterGroups,
  getVisibleDbHeaders,
  parseCrmWorkbook,
} from '../lib/excelParser';

describe('parseCrmWorkbook with test.xlsx', () => {
  const testXlsxPath = path.resolve(process.cwd(), 'test.xlsx');
  const fileBuffer = fs.readFileSync(testXlsxPath);

  test('parses all 3 worksheets (DB, 組別, 界別) and 50 DB headers', () => {
    const result = parseCrmWorkbook(fileBuffer, 'test.xlsx');

    expect(result.fileName).toBe('test.xlsx');
    expect(result.sheetNames).toEqual(['DB', '組別', '界別']);
    expect(result.dbHeaders.length).toBe(50);
    expect(result.dbHeaders[0]).toBe('酒會');
    expect(result.dbHeaders[1]).toBe('升旗');
    expect(result.dbHeaders[3]).toBe('組別');
    expect(result.dbHeaders[6]).toBe('姓名');
    expect(result.dbHeaders[48]).toBe('PID');
    expect(result.dbHeaders[49]).toBe('Pf.ID');
  });

  test('distinguishes DB section header rows from contact data rows and normalizes _x000D_ line breaks', () => {
    const result = parseCrmWorkbook(fileBuffer, 'test.xlsx');

    expect(result.dbRows.length).toBe(59);
    const sectionRows = result.dbRows.filter((r) => r.kind === 'section');
    const dataRows = result.dbRows.filter((r) => r.kind === 'record');

    expect(sectionRows.length).toBe(11);
    expect(dataRows.length).toBe(48);

    // Check first section row (Row 2: MA000 公共行政部門)
    expect(sectionRows[0]?.rowNumber).toBe(2);
    expect(sectionRows[0]?.sectionCode).toBe('MA000');
    expect(sectionRows[0]?.sectionTitle).toBe('公共行政部門');

    // Check Row 49 (MZ001 基本法委員會 with multiline Remark normalized)
    const mz001Section = sectionRows.find((r) => r.sectionCode === 'MZ001');
    expect(mz001Section).toBeDefined();
    expect(mz001Section?.sectionTitle).toBe('基本法委員會');
    expect(mz001Section?.values['Remark']).toContain('2024年12月31日更新\nhttps://appimg.modaily.cn');
    expect(mz001Section?.values['Remark']).not.toContain('_x000D_');
  });

  test('computes accurate attendance and category summary statistics', () => {
    const result = parseCrmWorkbook(fileBuffer, 'test.xlsx');

    expect(result.summary.totalRecords).toBe(48);
    expect(result.summary.sectionCount).toBe(11);
    expect(result.summary.attendingCount).toBe(39);
    expect(result.summary.notAttendingCount).toBe(9);
    expect(result.summary.receptionCount).toBe(39);
    expect(result.summary.flagRaisingCount).toBe(39);
    expect(result.summary.countedHeadcount).toBe(33);
    expect(result.summary.groupCount).toBe(558);
    expect(result.summary.sectorCount).toBe(26);
  });

  test('parses 組別 (558 rows) and 界別 (26 rows) and backfills Chinese names from DB section headers', () => {
    const result = parseCrmWorkbook(fileBuffer, 'test.xlsx');

    expect(result.groups.length).toBe(558);
    expect(result.sectors.length).toBe(26);

    const ma000Sector = result.sectors.find((s) => s.code === 'MA000');
    expect(ma000Sector?.nameZh).toBe('公共行政部門');

    const mf000Sector = result.sectors.find((s) => s.code === 'MF000');
    expect(mf000Sector?.remark).toBe('五年一屆');

    const mz001Group = result.groups.find((g) => g.code === 'MZ001');
    expect(mz001Group?.sectorCode).toBe('MZ000');
    expect(mz001Group?.category).toBe('委員會');
    expect(mz001Group?.nameZh).toBe('基本法委員會');
  });

  test('filters DB rows by search query, attendance status, and section visibility', () => {
    const result = parseCrmWorkbook(fileBuffer, 'test.xlsx');

    const onlyRecords = filterDbRows(result.dbRows, {
      searchQuery: '',
      attendanceFilter: 'all',
      showSectionHeaders: false,
    });
    expect(onlyRecords.length).toBe(48);

    const notAttending = filterDbRows(result.dbRows, {
      searchQuery: '',
      attendanceFilter: '不出席活動',
      showSectionHeaders: false,
    });
    expect(notAttending.length).toBe(9);

    const mz001Rows = filterDbRows(result.dbRows, {
      searchQuery: 'MZ001',
      attendanceFilter: 'all',
      showSectionHeaders: true,
    });
    // 1 MZ001 section header + 10 MZ001 data rows
    expect(mz001Rows.length).toBe(11);
  });

  test('filters 組別 rows by sector, category, and search query, and selects DB column presets', () => {
    const result = parseCrmWorkbook(fileBuffer, 'test.xlsx');

    const ma000Officials = filterGroups(result.groups, {
      searchQuery: '',
      sectorFilter: 'MA000',
      categoryFilter: '現任官員',
    });
    expect(ma000Officials.length).toBe(26);

    const coreHeaders = getVisibleDbHeaders(result.dbHeaders, 'core');
    expect(coreHeaders.length).toBeLessThan(result.dbHeaders.length);
    expect(coreHeaders).toContain('組別');
    expect(coreHeaders).toContain('出席活動');

    const allHeaders = getVisibleDbHeaders(result.dbHeaders, 'all');
    expect(allHeaders.length).toBe(50);
  });
});
