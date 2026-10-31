import { describe, expect, test } from 'bun:test';
import * as XLSX from 'xlsx';
import { filterDbRows, type DbRowItem } from '../lib/excelParser';
import {
  computeExternalFilterRowSet,
  getAvailableMatchKeys,
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
});
