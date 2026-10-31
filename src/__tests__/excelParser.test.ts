import { describe, expect, test } from 'bun:test';
import fs from 'node:fs';
import path from 'node:path';
import {
  computeDbFacetCounts,
  filterDbRows,
  filterGroups,
  getSectionRecordRowNumbers,
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
    // 2 ancestor section headers (MZ000 + MZ001) + 10 MZ001 data rows
    expect(mz001Rows.length).toBe(12);
    expect(mz001Rows[0].sectionCode).toBe('MZ000');
    expect(mz001Rows[1].sectionCode).toBe('MZ001');
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

  test('tracks 3-level ancestorSectionRows, sectorCode, and category on DB record rows', () => {
    const result = parseCrmWorkbook(fileBuffer, 'test.xlsx');
    const row5 = result.dbRows.find((r) => r.rowNumber === 5); // MA001-01 record
    expect(row5?.kind).toBe('record');
    expect(row5?.ancestorSectionRows).toEqual([2, 3, 4]); // MA000 (2), MA001 (3), MA001-01 (4)
    expect(row5?.sectorCode).toBe('MA000');
    expect(row5?.category).toBe('現任官員');

    const row16 = result.dbRows.find((r) => r.rowNumber === 16); // MA001-02 record
    expect(row16?.ancestorSectionRows).toEqual([2, 3, 15]); // MA000 (2), MA001 (3), MA001-02 (15)

    const row50 = result.dbRows.find((r) => r.rowNumber === 50); // MZ001 record
    expect(row50?.ancestorSectionRows).toEqual([48, 49]); // MZ000 (48), MZ001 (49)
    expect(row50?.sectorCode).toBe('MZ000');
    expect(row50?.category).toBe('委員會');
  });

  test('filters DB rows by sector, category, custom rules, selection state, and retains only matched ancestor sections', () => {
    const result = parseCrmWorkbook(fileBuffer, 'test.xlsx');

    const ma00102WithSections = filterDbRows(result.dbRows, {
      searchQuery: '',
      attendanceFilter: 'all',
      groupFilter: 'MA001-02',
      showSectionHeaders: true,
    });
    expect(ma00102WithSections.map((r) => r.rowNumber)).toEqual([2, 3, 15, 16]);

    const customFiltered = filterDbRows(result.dbRows, {
      searchQuery: '',
      attendanceFilter: 'all',
      showSectionHeaders: false,
      customRulesMode: 'AND',
      customRules: [
        { id: 'r1', column: '組別', operator: 'equals', value: 'MA002' },
        { id: 'r2', column: '計算', operator: 'equals', value: '0' },
      ],
    });
    expect(customFiltered.length).toBe(9);
    expect(customFiltered.every((r) => r.values['組別'] === 'MA002' && r.values['計算'] === '0')).toBe(true);

    const selectedSet = new Set<number>([5, 50]);
    const onlySelected = filterDbRows(result.dbRows, {
      searchQuery: '',
      attendanceFilter: 'all',
      selectionFilter: 'selected',
      selectedRowNumbers: selectedSet,
      showSectionHeaders: true,
    });
    expect(onlySelected.map((r) => r.rowNumber)).toEqual([2, 3, 4, 5, 48, 49, 50]);

    // Blank & special symbol filters must never include section rows (11 section rows have blank fields)
    const blankAttendance = filterDbRows(result.dbRows, {
      searchQuery: '',
      attendanceFilter: 'blank',
      showSectionHeaders: false,
    });
    expect(blankAttendance.length).toBe(0);

    const blankReception = filterDbRows(result.dbRows, {
      searchQuery: '',
      attendanceFilter: 'all',
      receptionFilter: 'blank',
      showSectionHeaders: false,
    });
    expect(blankReception.length).toBe(9);
    expect(blankReception.every((r) => r.kind === 'record' && (r.values['酒會'] ?? '') === '')).toBe(true);

    const receptionSymbol = filterDbRows(result.dbRows, {
      searchQuery: '',
      attendanceFilter: 'all',
      receptionFilter: '%',
      showSectionHeaders: false,
    });
    expect(receptionSymbol.length).toBe(39);
    expect(receptionSymbol.every((r) => r.kind === 'record' && r.values['酒會'] === '%')).toBe(true);

    const flagRaisingSymbol = filterDbRows(result.dbRows, {
      searchQuery: '',
      attendanceFilter: 'all',
      flagRaisingFilter: '*',
      showSectionHeaders: false,
    });
    expect(flagRaisingSymbol.length).toBe(39);
    expect(flagRaisingSymbol.every((r) => r.kind === 'record' && r.values['升旗'] === '*')).toBe(true);

    const emptyReceptionRule = filterDbRows(result.dbRows, {
      searchQuery: '',
      attendanceFilter: 'all',
      showSectionHeaders: false,
      customRules: [{ id: 'e1', column: '酒會', operator: 'is_empty', value: '' }],
    });
    expect(emptyReceptionRule.length).toBe(9);

    const nonEmptyReceptionRule = filterDbRows(result.dbRows, {
      searchQuery: '',
      attendanceFilter: 'all',
      showSectionHeaders: false,
      customRules: [{ id: 'ne1', column: '酒會', operator: 'is_not_empty', value: '' }],
    });
    expect(nonEmptyReceptionRule.length).toBe(39);
  });

  test('sorts DB record rows by column while hiding section headers during active sort', () => {
    const result = parseCrmWorkbook(fileBuffer, 'test.xlsx');
    const sortedDesc = filterDbRows(result.dbRows, {
      searchQuery: '',
      attendanceFilter: 'all',
      groupFilter: 'MZ001',
      showSectionHeaders: true,
      sort: { column: '出席活動', direction: 'desc' },
    });
    expect(sortedDesc.every((r) => r.kind === 'record')).toBe(true);
    // In MZ001, Row 50 is '不出席活動' and Row 51 is 'Y'; descending sort places 'Y' first and '不出席活動' last
    expect(sortedDesc[0]?.values['出席活動']).toBe('Y');
    expect(sortedDesc[sortedDesc.length - 1]?.values['出席活動']).toBe('不出席活動');

    const sortedAsc = filterDbRows(result.dbRows, {
      searchQuery: '',
      attendanceFilter: 'all',
      groupFilter: 'MZ001',
      showSectionHeaders: true,
      sort: { column: '出席活動', direction: 'asc' },
    });
    expect(sortedAsc[0]?.values['出席活動']).toBe('不出席活動');
    expect(sortedAsc[sortedAsc.length - 1]?.values['出席活動']).toBe('Y');
  });

  test('computes facet counts and section descendant record rowNumbers', () => {
    const result = parseCrmWorkbook(fileBuffer, 'test.xlsx');
    const counts = computeDbFacetCounts(result.dbRows);
    expect(counts.bySector.get('MA000')).toBe(38);
    expect(counts.bySector.get('MZ000')).toBe(10);
    expect(counts.byGroup.get('MZ001')).toBe(10);

    // Section row 3 (MA001) covers MA001-01 (10 rows), MA001-02 (1 row), MA001-03 (2 rows) = 13 rows
    const ma001Records = getSectionRecordRowNumbers(result.dbRows, 3);
    expect(ma001Records.length).toBe(13);
  });
});
