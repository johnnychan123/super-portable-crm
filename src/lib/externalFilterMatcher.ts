import * as XLSX from 'xlsx';
import type { DbRowItem } from './excelParser';

export interface ParsedExternalFilterSheet {
  fileName: string;
  sheetName: string;
  headers: string[];
  rows: Record<string, string>[];
  sharedHeaders: string[];
  extraHeaders: string[];
}

export interface ExternalMatchKeyOption {
  key: string;
  label: string;
  columns: string[];
}

export interface GroupMatchOverride {
  id: string;
  groupCode: string;
  matchKey: string;
}

export interface ExternalFilterMatchConfig {
  defaultMatchKey: string;
  groupOverrides?: ReadonlyArray<GroupMatchOverride>;
}

export interface ExternalFilterMatchResult {
  matchKey: string;
  resolvedKeyColumns: string[];
  matchedDbRowNumbers: Set<number>;
  matchedExternalRowCount: number;
  unmatchedExternalRowCount: number;
  totalExternalRowCount: number;
  extraValuesByDbRow: Map<number, Record<string, string>>;
}

export interface ImportedFilterFile {
  id: string;
  fileName: string;
  sheet: ParsedExternalFilterSheet;
  availableMatchKeys: ExternalMatchKeyOption[];
  selectedMatchKey: string;
  matchResult: ExternalFilterMatchResult;
}

export type ExternalFilterViewMode = 'all' | 'union' | 'intersection' | string;

const SECTION_CODE_REGEX = /^(M[A-Z]\d{3}(?:-\d{2})?)\s*(.*)$/;

function normalizeCellText(value: unknown): string {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value)
    .replace(/_x000D_\r*\n/g, '\n')
    .replace(/_x000D_/g, '\n')
    .replace(/\r+\n/g, '\n')
    .replace(/\r/g, '\n')
    .trim();
}

export function parseExternalFilterWorkbook(
  data: Uint8Array | ArrayBuffer,
  fileName: string,
  dbHeaders: string[],
): ParsedExternalFilterSheet {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const workbook = XLSX.read(bytes, { type: 'array', raw: false });
  const sheetNames = workbook.SheetNames;
  const sheetName = sheetNames.includes('DB') ? 'DB' : (sheetNames[0] ?? 'Sheet1');
  const sheet = sheetName ? workbook.Sheets[sheetName] : undefined;

  const headers: string[] = [];
  const rows: Record<string, string>[] = [];

  if (sheet) {
    const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: '',
      blankrows: false,
    });

    if (matrix.length > 0) {
      const headerRow = matrix[0] ?? [];
      const seenHeaders = new Set<string>();
      for (let i = 0; i < headerRow.length; i++) {
        const rawHeader = normalizeCellText(headerRow[i]);
        const header = rawHeader || `Col_${i + 1}`;
        if (!seenHeaders.has(header)) {
          seenHeaders.add(header);
          headers.push(header);
        } else {
          const dedupHeader = `${header}_${i + 1}`;
          seenHeaders.add(dedupHeader);
          headers.push(dedupHeader);
        }
      }

      for (let r = 1; r < matrix.length; r++) {
        const rawRow = matrix[r] ?? [];
        const values: Record<string, string> = {};
        let hasAnyValue = false;

        for (let c = 0; c < headers.length; c++) {
          const header = headers[c];
          const text = normalizeCellText(rawRow[c]);
          values[header] = text;
          if (text !== '') {
            hasAnyValue = true;
          }
        }

        if (!hasAnyValue) {
          continue;
        }

        // Detect and skip section header rows (e.g. MA000 公共行政部門)
        const groupVal = values['組別'] ?? '';
        const calcVal = values['計算'] ?? '';
        const attendVal = values['出席活動'] ?? '';
        const nameVal = values['姓名'] ?? '';
        const firstNonEmptyVal =
          nameVal || Object.values(values).find((v) => v !== '') || '';
        const nonEmptyCount = Object.values(values).filter((v) => v !== '').length;

        const firstColHeader = headers[0] ?? '';
        const firstColVal = firstColHeader ? (values[firstColHeader] ?? '') : '';
        const firstColMatch = firstColVal.match(SECTION_CODE_REGEX);
        const nonRemarkNonFirstEmpty = Object.entries(values).every(
          ([k, v]) => k === firstColHeader || k === 'Remark' || v === '',
        );
        const isFallbackSectionRow =
          !headers.includes('姓名') &&
          firstColMatch !== null &&
          ((firstColMatch[2] ?? '').trim() !== '' ||
            /^M[A-Z]000$/.test(firstColMatch[1] ?? '')) &&
          nonRemarkNonFirstEmpty;

        const isSectionRow =
          (groupVal === '' &&
            calcVal === '' &&
            attendVal === '' &&
            SECTION_CODE_REGEX.test(firstNonEmptyVal) &&
            (nameVal !== '' || nonEmptyCount <= 2)) ||
          isFallbackSectionRow;

        if (isSectionRow) {
          continue;
        }

        rows.push(values);
      }
    }
  }

  const dbHeaderSet = new Set(dbHeaders);
  const sharedHeaders = headers.filter((h) => dbHeaderSet.has(h));
  const extraHeaders = headers.filter((h) => !dbHeaderSet.has(h));

  return {
    fileName,
    sheetName,
    headers,
    rows,
    sharedHeaders,
    extraHeaders,
  };
}

const AUTO_CANDIDATE_KEYS: string[][] = [
  ['Pf.ID'],
  ['PID', '組別'],
  ['PID'],
  ['組別', '姓名'],
  ['姓名'],
  ['中文名(沒有稱謂)'],
];

function hasNonEmptyCompositeKeyInRows(
  rows: Record<string, string>[],
  cols: string[],
): boolean {
  if (cols.length === 0) return false;
  return rows.some((row) => cols.every((c) => (row[c] ?? '').trim() !== ''));
}

export function resolveAutoMatchKeyColumns(
  sheet: ParsedExternalFilterSheet,
  dbRows: DbRowItem[],
): string[] {
  const sharedSet = new Set(sheet.sharedHeaders);
  const recordDbRows = dbRows.filter((r) => r.kind === 'record').map((r) => r.values);

  for (const candidate of AUTO_CANDIDATE_KEYS) {
    if (
      candidate.every((col) => sharedSet.has(col)) &&
      hasNonEmptyCompositeKeyInRows(sheet.rows, candidate) &&
      hasNonEmptyCompositeKeyInRows(recordDbRows, candidate)
    ) {
      return candidate;
    }
  }

  for (const col of sheet.sharedHeaders) {
    if (
      hasNonEmptyCompositeKeyInRows(sheet.rows, [col]) &&
      hasNonEmptyCompositeKeyInRows(recordDbRows, [col])
    ) {
      return [col];
    }
  }

  return sheet.sharedHeaders.length > 0 ? [sheet.sharedHeaders[0]] : [];
}

export function getGlobalMatchKeyOptions(dbHeaders: string[]): ExternalMatchKeyOption[] {
  const headerSet = new Set(dbHeaders);
  const options: ExternalMatchKeyOption[] = [
    {
      key: 'auto',
      label: '自動偵測 (Pf.ID → PID+組別 → PID → 組別+姓名 → 姓名)',
      columns: [],
    },
  ];

  if (headerSet.has('PID') && headerSet.has('組別')) {
    options.push({
      key: 'PID+組別',
      label: '複合欄：PID + 組別',
      columns: ['PID', '組別'],
    });
  }
  if (headerSet.has('組別') && headerSet.has('姓名')) {
    options.push({
      key: '組別+姓名',
      label: '複合欄：組別 + 姓名',
      columns: ['組別', '姓名'],
    });
  }

  for (const col of dbHeaders) {
    options.push({
      key: col,
      label: `單一欄：${col}`,
      columns: [col],
    });
  }

  return options;
}

export function resolveMatchKeyColumns(
  matchKey: string,
  sheet: ParsedExternalFilterSheet,
  dbRows: DbRowItem[],
): string[] {
  const sharedSet = new Set(sheet.sharedHeaders);
  if (matchKey === 'auto') {
    return resolveAutoMatchKeyColumns(sheet, dbRows);
  }
  if (matchKey === 'PID+組別') {
    return sharedSet.has('PID') && sharedSet.has('組別') ? ['PID', '組別'] : [];
  }
  if (matchKey === '組別+姓名') {
    return sharedSet.has('組別') && sharedSet.has('姓名') ? ['組別', '姓名'] : [];
  }
  return sharedSet.has(matchKey) ? [matchKey] : [];
}

export function getAvailableMatchKeys(
  sheet: ParsedExternalFilterSheet,
  dbRows: DbRowItem[],
): ExternalMatchKeyOption[] {
  const autoCols = resolveAutoMatchKeyColumns(sheet, dbRows);
  const autoSuffix = autoCols.length > 0 ? ` (${autoCols.join(' + ')})` : '';
  const options: ExternalMatchKeyOption[] = [
    {
      key: 'global',
      label: '跟隨全域規則（預設 + 組別設定）',
      columns: autoCols,
    },
    {
      key: 'auto',
      label: `單檔自動偵測${autoSuffix}`,
      columns: autoCols,
    },
  ];

  const sharedSet = new Set(sheet.sharedHeaders);
  if (sharedSet.has('PID') && sharedSet.has('組別')) {
    options.push({
      key: 'PID+組別',
      label: '複合欄：PID + 組別',
      columns: ['PID', '組別'],
    });
  }
  if (sharedSet.has('組別') && sharedSet.has('姓名')) {
    options.push({
      key: '組別+姓名',
      label: '複合欄：組別 + 姓名',
      columns: ['組別', '姓名'],
    });
  }

  for (const col of sheet.sharedHeaders) {
    options.push({
      key: col,
      label: `單一欄：${col}`,
      columns: [col],
    });
  }

  return options;
}

function buildCompositeKey(values: Record<string, string>, columns: string[]): string | null {
  if (columns.length === 0) return null;
  const parts: string[] = [];
  for (const col of columns) {
    const val = (values[col] ?? '').trim().toLowerCase();
    if (val === '') {
      // Never allow blank cell values to participate in key collision
      return null;
    }
    parts.push(val);
  }
  return parts.join('\u0001');
}

export function matchExternalFilterToDb(
  sheet: ParsedExternalFilterSheet,
  dbRows: DbRowItem[],
  matchKey = 'auto',
  config?: ExternalFilterMatchConfig,
): ExternalFilterMatchResult {
  const useGlobalConfig = matchKey === 'global';
  const effectiveDefaultKey = useGlobalConfig
    ? (config?.defaultMatchKey ?? 'auto')
    : matchKey;
  const resolvedKeyColumns = resolveMatchKeyColumns(effectiveDefaultKey, sheet, dbRows);

  const groupOverrideColumnsMap = new Map<string, string[]>();
  if (useGlobalConfig && config?.groupOverrides) {
    for (const override of config.groupOverrides) {
      const normalizedGroup = override.groupCode.trim().toUpperCase();
      if (!normalizedGroup) continue;
      groupOverrideColumnsMap.set(
        normalizedGroup,
        resolveMatchKeyColumns(override.matchKey, sheet, dbRows),
      );
    }
  }

  const getColumnsForRow = (values: Record<string, string>): string[] => {
    if (groupOverrideColumnsMap.size > 0) {
      const grp = (values['組別'] ?? '').trim().toUpperCase();
      if (grp && groupOverrideColumnsMap.has(grp)) {
        return groupOverrideColumnsMap.get(grp) ?? [];
      }
    }
    return resolvedKeyColumns;
  };

  const getNamespacedKey = (values: Record<string, string>): string | null => {
    const cols = getColumnsForRow(values);
    const composite = buildCompositeKey(values, cols);
    if (!composite) return null;
    return `${cols.join('+')}::${composite}`;
  };

  const dbRecordsByKey = new Map<string, number[]>();
  for (const row of dbRows) {
    if (row.kind !== 'record') continue;
    const key = getNamespacedKey(row.values);
    if (!key) continue;
    const list = dbRecordsByKey.get(key) ?? [];
    list.push(row.rowNumber);
    dbRecordsByKey.set(key, list);
  }

  const matchedDbRowNumbers = new Set<number>();
  const extraValuesByDbRow = new Map<number, Record<string, string>>();
  let matchedExternalRowCount = 0;
  let unmatchedExternalRowCount = 0;

  for (const extRow of sheet.rows) {
    const key = getNamespacedKey(extRow);
    const matchedRowNums = key ? dbRecordsByKey.get(key) : undefined;
    if (matchedRowNums && matchedRowNums.length > 0) {
      matchedExternalRowCount++;
      const extraRecord: Record<string, string> = {};
      for (const extraCol of sheet.extraHeaders) {
        extraRecord[extraCol] = extRow[extraCol] ?? '';
      }
      for (const rowNum of matchedRowNums) {
        matchedDbRowNumbers.add(rowNum);
        if (sheet.extraHeaders.length > 0) {
          const prev = extraValuesByDbRow.get(rowNum) ?? {};
          extraValuesByDbRow.set(rowNum, { ...prev, ...extraRecord });
        }
      }
    } else {
      unmatchedExternalRowCount++;
    }
  }

  return {
    matchKey,
    resolvedKeyColumns,
    matchedDbRowNumbers,
    matchedExternalRowCount,
    unmatchedExternalRowCount,
    totalExternalRowCount: sheet.rows.length,
    extraValuesByDbRow,
  };
}

export function computeExternalFilterRowSet(
  files: ImportedFilterFile[],
  mode: ExternalFilterViewMode,
): Set<number> | null {
  if (mode === 'all' || files.length === 0) {
    return null;
  }

  if (mode === 'union') {
    const unionSet = new Set<number>();
    for (const f of files) {
      for (const r of f.matchResult.matchedDbRowNumbers) {
        unionSet.add(r);
      }
    }
    return unionSet;
  }

  if (mode === 'intersection') {
    const [first, ...rest] = files;
    if (!first) return new Set<number>();
    const interSet = new Set<number>(first.matchResult.matchedDbRowNumbers);
    for (const f of rest) {
      for (const r of Array.from(interSet)) {
        if (!f.matchResult.matchedDbRowNumbers.has(r)) {
          interSet.delete(r);
        }
      }
    }
    return interSet;
  }

  const targetFile = files.find((f) => f.id === mode);
  if (!targetFile) {
    return null;
  }
  return new Set<number>(targetFile.matchResult.matchedDbRowNumbers);
}

export function mergeExtraValuesFromFiles(
  files: ImportedFilterFile[],
): Map<number, Record<string, string>> {
  const merged = new Map<number, Record<string, string>>();
  for (const f of files) {
    for (const [rowNum, extraVals] of f.matchResult.extraValuesByDbRow.entries()) {
      const current = merged.get(rowNum) ?? {};
      const next: Record<string, string> = { ...current };
      for (const [col, val] of Object.entries(extraVals)) {
        if (val !== '' || !(col in next)) {
          next[col] = val;
        }
      }
      merged.set(rowNum, next);
    }
  }
  return merged;
}
