import { describe, expect, test } from 'bun:test';
import * as XLSX from 'xlsx';
import { filterDbRows, type DbRowItem } from '../lib/excelParser';
import {
  applyPrefixGroupSelection,
  computeExternalFilterRowSet,
  filterGroupsByPrefix,
  getAvailableMatchKeys,
  getGroupOptionsForOverride,
  matchExternalFilterToDb,
  mergeExtraValuesFromFiles,
  parseExternalFilterWorkbook,
} from '../lib/externalFilterMatcher';

function makeWorkbookBuffer(sheetName: string, aoa: string[][]): Uint8Array {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  return new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer);
}

describe('externalFilterMatcher', () => {
  const dbHeaders = ['組別', '計算', '出席活動', '姓名', '職位', 'PID', 'Pf.ID'];
  const dbRows: DbRowItem[] = [
    {
      rowNumber: 2,
      kind: 'section',
      sectionCode: 'MA000',
      sectionTitle: '公共行政部門',
      ancestorSectionRows: [],
      values: {
        組別: '',
        計算: '',
        出席活動: '',
        姓名: 'MA000 公共行政部門',
        職位: '',
        PID: '',
        'Pf.ID': '',
      },
    },
    {
      rowNumber: 3,
      kind: 'record',
      ancestorSectionRows: [2],
      values: {
        組別: 'MA001',
        計算: '1',
        出席活動: 'Y',
        姓名: '陳大文',
        職位: '處長',
        PID: 'P001',
        'Pf.ID': 'PF-101',
      },
    },
    {
      rowNumber: 4,
      kind: 'record',
      ancestorSectionRows: [2],
      values: {
        組別: 'MA001',
        計算: '1',
        出席活動: 'Y',
        姓名: '李小明',
        職位: '主任',
        PID: 'P002',
        'Pf.ID': '',
      },
    },
    {
      rowNumber: 5,
      kind: 'record',
      ancestorSectionRows: [2],
      values: {
        組別: 'MA002',
        計算: '0',
        出席活動: '不出席活動',
        姓名: '張三',
        職位: '顧問',
        PID: '',
        'Pf.ID': '',
      },
    },
  ];

  test('parses modified past-export Excel (removed DB columns + added event columns) and skips section rows', () => {
    const buf = makeWorkbookBuffer('Sheet1', [
      ['組別', '姓名', '桌號', '車牌'],
      ['', 'MA000 公共行政部門', '', ''], // section row - should be skipped
      ['MA001', '陳大文', 'A1', 'MA-12-34'],
      ['MA002', '王五(不在DB)', 'B2', ''],
    ]);

    const parsed = parseExternalFilterWorkbook(buf, 'past_event.xlsx', dbHeaders);
    expect(parsed.sharedHeaders).toEqual(['組別', '姓名']);
    expect(parsed.extraHeaders).toEqual(['桌號', '車牌']);
    expect(parsed.rows.length).toBe(2);

    const match = matchExternalFilterToDb(parsed, dbRows, 'auto');
    expect(match.resolvedKeyColumns).toEqual(['組別', '姓名']);
    expect(Array.from(match.matchedDbRowNumbers)).toEqual([3]);
    expect(match.matchedExternalRowCount).toBe(1);
    expect(match.unmatchedExternalRowCount).toBe(1);
    expect(match.extraValuesByDbRow.get(3)).toEqual({ 桌號: 'A1', 車牌: 'MA-12-34' });
  });

  test('never matches empty cells against empty cells and supports switching match keys', () => {
    const buf = makeWorkbookBuffer('DB', [
      ['PID', '姓名', '桌號'],
      ['', '張三', 'C3'],
    ]);

    const parsed = parseExternalFilterWorkbook(buf, 'pid_missing.xlsx', dbHeaders);
    const byPid = matchExternalFilterToDb(parsed, dbRows, 'PID');
    expect(byPid.matchedDbRowNumbers.size).toBe(0);

    const byName = matchExternalFilterToDb(parsed, dbRows, '姓名');
    expect(Array.from(byName.matchedDbRowNumbers)).toEqual([5]);
    expect(byName.extraValuesByDbRow.get(5)).toEqual({ 桌號: 'C3' });
  });

  test('combines multiple imported filter files with union, intersection, and single-file filter modes', () => {
    const buf1 = makeWorkbookBuffer('DB', [
      ['組別', '姓名', '桌號'],
      ['MA001', '陳大文', 'T1'],
      ['MA001', '李小明', 'T2'],
    ]);
    const buf2 = makeWorkbookBuffer('DB', [
      ['組別', '姓名', '車牌'],
      ['MA001', '李小明', 'CAR-99'],
      ['MA002', '張三', 'CAR-88'],
    ]);

    const s1 = parseExternalFilterWorkbook(buf1, 'f1.xlsx', dbHeaders);
    const s2 = parseExternalFilterWorkbook(buf2, 'f2.xlsx', dbHeaders);
    const files = [
      {
        id: 'f1',
        fileName: 'f1.xlsx',
        sheet: s1,
        availableMatchKeys: getAvailableMatchKeys(s1, dbRows),
        selectedMatchKey: 'auto',
        matchResult: matchExternalFilterToDb(s1, dbRows, 'auto'),
      },
      {
        id: 'f2',
        fileName: 'f2.xlsx',
        sheet: s2,
        availableMatchKeys: getAvailableMatchKeys(s2, dbRows),
        selectedMatchKey: 'auto',
        matchResult: matchExternalFilterToDb(s2, dbRows, 'auto'),
      },
    ];

    expect(computeExternalFilterRowSet(files, 'all')).toBeNull();
    expect(Array.from(computeExternalFilterRowSet(files, 'union') ?? [])).toEqual([3, 4, 5]);
    expect(Array.from(computeExternalFilterRowSet(files, 'intersection') ?? [])).toEqual([4]);
    expect(Array.from(computeExternalFilterRowSet(files, 'f1') ?? [])).toEqual([3, 4]);

    const mergedExtra = mergeExtraValuesFromFiles(files);
    expect(mergedExtra.get(4)).toEqual({ 桌號: 'T2', 車牌: 'CAR-99' });

    const filteredWithSections = filterDbRows(dbRows, {
      searchQuery: '',
      attendanceFilter: 'all',
      externalMatchedRowNumbers: computeExternalFilterRowSet(files, 'intersection'),
      showSectionHeaders: true,
    });
    expect(filteredWithSections.map((r) => r.rowNumber)).toEqual([2, 4]);
  });

  test('skips fallback section rows when re-importing a custom export that omitted 姓名', () => {
    const buf = makeWorkbookBuffer('DB', [
      ['組別', 'PID', '出席活動', 'Remark'],
      ['MA000 公共行政部門', '', '', ''],
      ['MZ001 基本法委員會', '', '', '2024年12月31日更新'],
      ['MA001', '101', '1', ''],
    ]);

    const parsed = parseExternalFilterWorkbook(buf, 'custom_no_name.xlsx', dbHeaders);
    expect(parsed.rows.length).toBe(1);
    expect(parsed.rows[0]?.['PID']).toBe('101');
  });

  test('supports defaultMatchKey plus per-group matchKey overrides with strict blank non-fallback', () => {
    const buf = makeWorkbookBuffer('DB', [
      ['組別', 'PID', '姓名', '桌號'],
      ['MA001', 'P001', '陳大文', 'A1'], // MA001 uses default (PID+組別) -> matches row 3
      ['MA001', '', '李小明', 'A2'], // MA001 missing PID -> strict no-fallback -> should NOT match row 4
      ['MZ001', '', '張三', 'Z1'], // MZ001 overridden to use 組別+姓名 -> matches row 5 even without PID
    ]);

    const extendedDbRows: DbRowItem[] = [
      ...dbRows.slice(0, 3),
      {
        rowNumber: 5,
        kind: 'record',
        values: {
          'Pf.ID': '',
          PID: '',
          組別: 'MZ001',
          姓名: '張三',
          '中文名(沒有稱謂)': '張三',
        },
        ancestorSectionRows: [2],
      },
    ];

    const parsed = parseExternalFilterWorkbook(buf, 'mixed_groups.xlsx', dbHeaders);
    const res = matchExternalFilterToDb(parsed, extendedDbRows, 'global', {
      defaultMatchKey: 'PID+組別',
      groupOverrides: [{ id: 'g1', groupCode: 'MZ001', matchKey: '組別+姓名' }],
    });

    expect(Array.from(res.matchedDbRowNumbers)).toEqual([3, 5]);
    expect(res.matchedExternalRowCount).toBe(2);
    expect(res.unmatchedExternalRowCount).toBe(1);
    expect(res.extraValuesByDbRow.get(3)).toEqual({ 桌號: 'A1' });
    expect(res.extraValuesByDbRow.get(5)).toEqual({ 桌號: 'Z1' });

    // If a group override specifies a column omitted from the external sheet (e.g., Pf.ID), it safely yields 0 matches for that group
    const resMissingCol = matchExternalFilterToDb(parsed, extendedDbRows, 'global', {
      defaultMatchKey: 'PID+組別',
      groupOverrides: [{ id: 'g1', groupCode: 'MZ001', matchKey: 'Pf.ID' }],
    });
    expect(Array.from(resMissingCol.matchedDbRowNumbers)).toEqual([3]);

    // getGroupOptionsForOverride prioritizes groups appearing in external files and includes DB/external sub-groups not in workbook.groups
    const groupOpts = getGroupOptionsForOverride(
      [
        { id: '1', sectorCode: 'MA000', category: 'A', code: 'MA009', nameZh: '未出現部門' },
        { id: '2', sectorCode: 'MA000', category: 'A', code: 'MA001', nameZh: '行政長官辦公室' },
      ],
      extendedDbRows,
      [
        {
          id: 'f1',
          fileName: 'mixed_groups.xlsx',
          sheet: parsed,
          availableMatchKeys: [],
          selectedMatchKey: 'global',
          matchResult: res,
        },
      ],
    );
    expect(groupOpts.map((g) => g.code)).toEqual(['MA001', 'MZ001', 'MA009']);
    expect(groupOpts[0]?.externalCount).toBe(2);
    expect(groupOpts[1]?.externalCount).toBe(1);
    expect(groupOpts[2]?.externalCount).toBe(0);
  });

  test('supports multi-group groupCodes in a single rule and prefix batch selection/deselection', () => {
    const buf = makeWorkbookBuffer('DB', [
      ['組別', 'PID', '姓名', '桌號'],
      ['MA001', '', '陳大文', 'A1'], // MA001 uses default PID+組別 -> missing PID -> not matched
      ['MZ001', '', '李小明', 'Z1'], // MZ001 in multi-group rule (組別+姓名) -> matched
      ['MZ002', '', '張三', 'Z2'], // MZ002 in multi-group rule (組別+姓名) -> matched
    ]);

    const multiGroupDbRows: DbRowItem[] = [
      {
        rowNumber: 3,
        kind: 'record',
        values: { 組別: 'MA001', PID: 'P001', 姓名: '陳大文' },
        ancestorSectionRows: [],
      },
      {
        rowNumber: 4,
        kind: 'record',
        values: { 組別: 'MZ001', PID: '', 姓名: '李小明' },
        ancestorSectionRows: [],
      },
      {
        rowNumber: 5,
        kind: 'record',
        values: { 組別: 'MZ002', PID: '', 姓名: '張三' },
        ancestorSectionRows: [],
      },
    ];

    const groupOpts = [
      { code: 'MA001', nameZh: '行政長官辦公室', category: 'A', externalCount: 1 },
      { code: 'MZ001', nameZh: '基本法委員會', category: 'B', externalCount: 1 },
      { code: 'MZ002', nameZh: '選舉委員會', category: 'B', externalCount: 1 },
    ];

    expect(filterGroupsByPrefix(groupOpts, 'mz').map((g) => g.code)).toEqual([
      'MZ001',
      'MZ002',
    ]);

    const selectedAfterPrefix = applyPrefixGroupSelection(
      ['MA009'],
      groupOpts,
      'MZ',
      'select',
    );
    expect(selectedAfterPrefix).toEqual(['MA009', 'MZ001', 'MZ002']);

    const selectedAfterDeselect = applyPrefixGroupSelection(
      selectedAfterPrefix,
      groupOpts,
      'MZ001',
      'deselect',
    );
    expect(selectedAfterDeselect).toEqual(['MA009', 'MZ002']);

    const parsed = parseExternalFilterWorkbook(buf, 'multi_mz.xlsx', dbHeaders);
    const res = matchExternalFilterToDb(parsed, multiGroupDbRows, 'global', {
      defaultMatchKey: 'PID+組別',
      groupOverrides: [
        { id: 'rule-mz', groupCodes: ['MZ001', 'MZ002'], matchKey: '組別+姓名' },
      ],
    });

    expect(Array.from(res.matchedDbRowNumbers)).toEqual([4, 5]);
  });
});
