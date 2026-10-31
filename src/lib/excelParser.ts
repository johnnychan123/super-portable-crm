import * as XLSX from 'xlsx';

export interface DbRowItem {
  rowNumber: number;
  kind: 'section' | 'record';
  sectionCode?: string;
  sectionTitle?: string;
  currentSectionLabel?: string;
  ancestorSectionRows: number[];
  sectorCode?: string;
  category?: string;
  groupNameZh?: string;
  searchText?: string;
  values: Record<string, string>;
}

export interface CrmGroupItem {
  id: string;
  sectorCode: string;
  category: string;
  code: string;
  nameZh: string;
}

export interface CrmSectorItem {
  id: string;
  code: string;
  nameZh: string;
  remark: string;
}

export interface CrmWorkbookSummary {
  totalRecords: number;
  sectionCount: number;
  attendingCount: number;
  notAttendingCount: number;
  receptionCount: number;
  flagRaisingCount: number;
  countedHeadcount: number;
  groupCount: number;
  sectorCount: number;
}

export interface ParsedCrmWorkbook {
  fileName: string;
  filePath?: string;
  sheetNames: string[];
  dbHeaders: string[];
  dbRows: DbRowItem[];
  groups: CrmGroupItem[];
  sectors: CrmSectorItem[];
  rawGroupRows: string[][];
  rawSectorRows: string[][];
  summary: CrmWorkbookSummary;
}

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

const SECTION_CODE_REGEX = /^(M[A-Z]\d{3}(?:-\d{2})?)\s*(.*)$/;

export function parseCrmWorkbook(
  data: Uint8Array | ArrayBuffer,
  fileName = 'workbook.xlsx',
  filePath?: string,
): ParsedCrmWorkbook {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const workbook = XLSX.read(bytes, { type: 'array', raw: false });
  const sheetNames = workbook.SheetNames;

  const dbSheetName = sheetNames.includes('DB') ? 'DB' : sheetNames[0];
  const dbSheet = dbSheetName ? workbook.Sheets[dbSheetName] : undefined;

  const dbHeaders: string[] = [];
  const dbRows: DbRowItem[] = [];
  const sectionNameMap = new Map<string, string>();

  if (dbSheet) {
    const matrix = XLSX.utils.sheet_to_json<unknown[]>(dbSheet, {
      header: 1,
      defval: '',
      blankrows: false,
    });

    if (matrix.length > 0) {
      const headerRow = matrix[0] ?? [];
      for (let i = 0; i < headerRow.length; i++) {
        const header = normalizeCellText(headerRow[i]);
        dbHeaders.push(header || `Col_${i + 1}`);
      }

      let currentSectionLabel = '';
      let activeLevel0: { rowNumber: number; code: string } | null = null;
      let activeLevel1: { rowNumber: number; code: string } | null = null;
      let activeLevel2: { rowNumber: number; code: string } | null = null;

      for (let r = 1; r < matrix.length; r++) {
        const rawRow = matrix[r] ?? [];
        const values: Record<string, string> = {};
        let hasAnyValue = false;

        for (let c = 0; c < dbHeaders.length; c++) {
          const header = dbHeaders[c];
          const text = normalizeCellText(rawRow[c]);
          values[header] = text;
          if (text !== '') {
            hasAnyValue = true;
          }
        }

        if (!hasAnyValue) {
          continue;
        }

        const rowNumber = r + 1;
        const groupVal = values['組別'] ?? '';
        const nameVal = values['姓名'] ?? '';
        const match = groupVal === '' ? SECTION_CODE_REGEX.exec(nameVal) : null;

        if (match) {
          const sectionCode = match[1];
          const sectionTitle = match[2]?.trim() ?? '';
          currentSectionLabel = nameVal;
          if (sectionTitle && !/^X+$/i.test(sectionTitle)) {
            sectionNameMap.set(sectionCode, sectionTitle);
          }

          let ancestorSectionRows: number[] = [];
          if (sectionCode.endsWith('000')) {
            activeLevel0 = { rowNumber, code: sectionCode };
            activeLevel1 = null;
            activeLevel2 = null;
            ancestorSectionRows = [];
          } else if (!sectionCode.includes('-')) {
            if (activeLevel0 && sectionCode.slice(0, 2) !== activeLevel0.code.slice(0, 2)) {
              activeLevel0 = null;
            }
            activeLevel1 = { rowNumber, code: sectionCode };
            activeLevel2 = null;
            ancestorSectionRows = activeLevel0 ? [activeLevel0.rowNumber] : [];
          } else {
            if (activeLevel0 && sectionCode.slice(0, 2) !== activeLevel0.code.slice(0, 2)) {
              activeLevel0 = null;
            }
            if (activeLevel1 && !sectionCode.startsWith(`${activeLevel1.code}-`)) {
              activeLevel1 = null;
            }
            activeLevel2 = { rowNumber, code: sectionCode };
            ancestorSectionRows = [
              ...(activeLevel0 ? [activeLevel0.rowNumber] : []),
              ...(activeLevel1 ? [activeLevel1.rowNumber] : []),
            ];
          }

          dbRows.push({
            rowNumber,
            kind: 'section',
            sectionCode,
            sectionTitle,
            currentSectionLabel,
            ancestorSectionRows,
            values,
          });
        } else {
          const ancestorSectionRows = [
            ...(activeLevel0 ? [activeLevel0.rowNumber] : []),
            ...(activeLevel1 ? [activeLevel1.rowNumber] : []),
            ...(activeLevel2 ? [activeLevel2.rowNumber] : []),
          ];
          dbRows.push({
            rowNumber,
            kind: 'record',
            currentSectionLabel,
            ancestorSectionRows,
            values,
          });
        }
      }
    }
  }

  // Parse 組別 sheet
  const groups: CrmGroupItem[] = [];
  const rawGroupRows: string[][] = [];
  const groupSheet = workbook.Sheets['組別'];
  if (groupSheet) {
    const groupMatrix = XLSX.utils.sheet_to_json<unknown[]>(groupSheet, {
      header: 1,
      defval: '',
      blankrows: false,
    });
    for (let r = 0; r < groupMatrix.length; r++) {
      const row = (groupMatrix[r] ?? []).map((cell) => normalizeCellText(cell));
      rawGroupRows.push(row);
      if (r === 0) {
        continue;
      }
      const id = row[0] ?? '';
      const sectorCode = row[1] ?? '';
      const category = row[2] ?? '';
      const code = row[3] ?? '';
      const rawNameZh = row[4] ?? '';
      if (!id && !code) {
        continue;
      }
      groups.push({
        id,
        sectorCode,
        category,
        code,
        nameZh: rawNameZh || sectionNameMap.get(code) || '',
      });
    }
  }

  // Parse 界別 sheet
  const sectors: CrmSectorItem[] = [];
  const rawSectorRows: string[][] = [];
  const sectorSheet = workbook.Sheets['界別'];
  if (sectorSheet) {
    const sectorMatrix = XLSX.utils.sheet_to_json<unknown[]>(sectorSheet, {
      header: 1,
      defval: '',
      blankrows: false,
    });
    for (let r = 0; r < sectorMatrix.length; r++) {
      const row = (sectorMatrix[r] ?? []).map((cell) => normalizeCellText(cell));
      rawSectorRows.push(row);
      if (r === 0) {
        continue;
      }
      const id = row[0] ?? '';
      const code = row[1] ?? '';
      const rawNameZh = row[2] ?? '';
      const remark = row[3] ?? '';
      if (!id && !code) {
        continue;
      }
      sectors.push({
        id,
        code,
        nameZh: rawNameZh || sectionNameMap.get(code) || '',
        remark,
      });
    }
  }

  // Enrich DB rows with group/sector metadata and precomputed searchText
  const groupByCode = new Map<string, CrmGroupItem>();
  for (const g of groups) {
    groupByCode.set(g.code, g);
    if (g.code.endsWith('-00')) {
      const baseCode = g.code.slice(0, -3);
      if (!groupByCode.has(baseCode)) {
        groupByCode.set(baseCode, g);
      }
    }
  }

  for (const row of dbRows) {
    if (row.kind === 'section') {
      const secCode = row.sectionCode ?? '';
      const matchedGroup = groupByCode.get(secCode);
      row.sectorCode =
        matchedGroup?.sectorCode || (secCode.length >= 2 ? `${secCode.slice(0, 2)}000` : '');
      row.category = matchedGroup?.category || '';
      row.groupNameZh = row.sectionTitle || matchedGroup?.nameZh || '';
      row.searchText = [
        secCode,
        row.sectionTitle ?? '',
        row.sectorCode,
        row.category,
        ...Object.values(row.values),
      ]
        .join(' ')
        .toLowerCase();
    } else {
      const groupCode = row.values['組別'] ?? '';
      const matchedGroup = groupByCode.get(groupCode);
      row.sectorCode =
        matchedGroup?.sectorCode || (groupCode.length >= 2 ? `${groupCode.slice(0, 2)}000` : '');
      row.category = matchedGroup?.category || '';
      row.groupNameZh = matchedGroup?.nameZh || '';
      row.searchText = [
        row.currentSectionLabel ?? '',
        row.sectorCode,
        row.category,
        row.groupNameZh,
        ...Object.values(row.values),
      ]
        .join(' ')
        .toLowerCase();
    }
  }

  const recordRows = dbRows.filter((row) => row.kind === 'record');
  const sectionRows = dbRows.filter((row) => row.kind === 'section');

  const summary: CrmWorkbookSummary = {
    totalRecords: recordRows.length,
    sectionCount: sectionRows.length,
    attendingCount: recordRows.filter((r) => r.values['出席活動'] === 'Y').length,
    notAttendingCount: recordRows.filter((r) => r.values['出席活動'] === '不出席活動').length,
    receptionCount: recordRows.filter((r) => r.values['酒會'] === '%').length,
    flagRaisingCount: recordRows.filter((r) => r.values['升旗'] === '*').length,
    countedHeadcount: recordRows.filter((r) => r.values['計算'] === '1').length,
    groupCount: groups.length,
    sectorCount: sectors.length,
  };

  return {
    fileName,
    filePath,
    sheetNames,
    dbHeaders,
    dbRows,
    groups,
    sectors,
    rawGroupRows,
    rawSectorRows,
    summary,
  };
}

export type DbColumnPreset = 'core' | 'contact' | 'personal' | 'all';

const PRESET_COLUMNS: Record<Exclude<DbColumnPreset, 'all'>, string[]> = {
  core: [
    '酒會',
    '升旗',
    '計算',
    '組別',
    '出席活動',
    'Remark',
    '姓名',
    '職位',
    '部門/機構',
    '電話',
    '手機 (資料不對外使用)',
    'SMS',
    '不出席原因',
    'PID',
    'Pf.ID',
  ],
  contact: [
    '組別',
    '姓名',
    '職位',
    '部門/機構',
    '電話',
    '電話 (關聯資料)',
    '手機 (資料不對外使用)',
    'SMS',
    '最近一次大型活動 SMS',
    '傳真',
    '地址',
    '組別地址',
    'Label抬頭',
    '地區',
    '電郵(個人) 1',
    '電郵(個人) 2',
    '個人地址(中)',
    '手機(個人) 1',
    '已納入通訊錄',
  ],
  personal: [
    '組別',
    '姓名',
    '中文名(沒有稱謂)',
    '外文名(沒有稱謂)',
    '外文名',
    '外文職位',
    '出生日期',
    '年齡',
    '身份證資料 (只供內部使用)',
    '證件到期日',
    '配偶姓名(中)',
    '配偶姓名(葡)',
    '語言',
    '婚姻狀態',
    '邀請伉儷(個人)',
    '回鄉證編號',
    '回鄉證到期日',
    'PID',
    'Pf.ID',
  ],
};

export function getVisibleDbHeaders(allHeaders: string[], preset: DbColumnPreset): string[] {
  if (preset === 'all') {
    return allHeaders;
  }
  const targetSet = new Set(PRESET_COLUMNS[preset]);
  const filtered = allHeaders.filter((h) => targetSet.has(h));
  return filtered.length > 0 ? filtered : allHeaders;
}

export type CustomFilterOperator =
  | 'contains'
  | 'not_contains'
  | 'equals'
  | 'not_equals'
  | 'is_empty'
  | 'is_not_empty';

export interface CustomFilterRule {
  id: string;
  column: string;
  operator: CustomFilterOperator;
  value: string;
}

export interface DbSortOption {
  column: string | null;
  direction: 'asc' | 'desc';
}

export interface DbFilterOptions {
  searchQuery: string;
  attendanceFilter: 'all' | 'Y' | '不出席活動' | 'blank';
  receptionFilter?: 'all' | '%' | 'blank';
  flagRaisingFilter?: 'all' | '*' | 'blank';
  calcFilter?: 'all' | '1' | '0';
  sectorFilter?: string;
  categoryFilter?: string;
  groupFilter?: string;
  selectionFilter?: 'all' | 'selected' | 'unselected';
  selectedRowNumbers?: ReadonlySet<number>;
  externalMatchedRowNumbers?: ReadonlySet<number> | null;
  customRules?: CustomFilterRule[];
  customRulesMode?: 'AND' | 'OR';
  showSectionHeaders: boolean;
  sort?: DbSortOption;
}

const zhCollator = new Intl.Collator('zh-Hant', {
  numeric: true,
  sensitivity: 'base',
});

function matchesSearchQuery(haystack: string, rawQuery: string): boolean {
  const trimmed = rawQuery.trim().toLowerCase();
  if (!trimmed) {
    return true;
  }
  const orGroups = trimmed
    .split('|')
    .map((g) => g.trim())
    .filter(Boolean);
  if (orGroups.length === 0) {
    return true;
  }
  return orGroups.some((group) => {
    const andTokens = group.split(/\s+/).filter(Boolean);
    return andTokens.every((token) => haystack.includes(token));
  });
}

function evaluateCustomRule(row: DbRowItem, rule: CustomFilterRule): boolean {
  const cellVal = (row.values[rule.column] ?? '').trim();
  const targetVal = rule.value.trim();

  switch (rule.operator) {
    case 'is_empty':
      return cellVal === '';
    case 'is_not_empty':
      return cellVal !== '';
    case 'contains':
      return targetVal === '' ? true : cellVal.toLowerCase().includes(targetVal.toLowerCase());
    case 'not_contains':
      return targetVal === '' ? true : !cellVal.toLowerCase().includes(targetVal.toLowerCase());
    case 'equals':
      return targetVal === '' ? true : cellVal.toLowerCase() === targetVal.toLowerCase();
    case 'not_equals':
      return targetVal === '' ? true : cellVal.toLowerCase() !== targetVal.toLowerCase();
    default:
      return true;
  }
}

function isRuleEffective(rule: CustomFilterRule): boolean {
  if (rule.operator === 'is_empty' || rule.operator === 'is_not_empty') {
    return true;
  }
  return rule.value.trim() !== '';
}

export function filterDbRows(rows: DbRowItem[], options: DbFilterOptions): DbRowItem[] {
  const query = options.searchQuery.trim();
  const attendanceFilter = options.attendanceFilter ?? 'all';
  const receptionFilter = options.receptionFilter ?? 'all';
  const flagRaisingFilter = options.flagRaisingFilter ?? 'all';
  const calcFilter = options.calcFilter ?? 'all';
  const sectorFilter = options.sectorFilter ?? 'all';
  const categoryFilter = options.categoryFilter ?? 'all';
  const groupFilter = options.groupFilter ?? 'all';
  const selectionFilter = options.selectionFilter ?? 'all';
  const selectedSet = options.selectedRowNumbers;
  const externalMatchedSet = options.externalMatchedRowNumbers ?? null;
  const effectiveRules = (options.customRules ?? []).filter(isRuleEffective);
  const rulesMode = options.customRulesMode ?? 'AND';
  const sortColumn = options.sort?.column ?? null;
  const sortDirection = options.sort?.direction ?? 'asc';

  const hasNonSearchFilter =
    attendanceFilter !== 'all' ||
    receptionFilter !== 'all' ||
    flagRaisingFilter !== 'all' ||
    calcFilter !== 'all' ||
    sectorFilter !== 'all' ||
    categoryFilter !== 'all' ||
    groupFilter !== 'all' ||
    selectionFilter !== 'all' ||
    externalMatchedSet !== null ||
    effectiveRules.length > 0;

  const hasAnyFilter = query !== '' || hasNonSearchFilter;

  const matchedRecordRowNumbers = new Set<number>();
  const matchedSectionRowNumbers = new Set<number>();

  for (const row of rows) {
    if (row.kind !== 'record') {
      continue;
    }

    if (attendanceFilter !== 'all') {
      const val = row.values['出席活動'] ?? '';
      if (attendanceFilter === 'blank' ? val !== '' : val !== attendanceFilter) {
        continue;
      }
    }

    if (receptionFilter !== 'all') {
      const val = row.values['酒會'] ?? '';
      if (receptionFilter === 'blank' ? val !== '' : val !== receptionFilter) {
        continue;
      }
    }

    if (flagRaisingFilter !== 'all') {
      const val = row.values['升旗'] ?? '';
      if (flagRaisingFilter === 'blank' ? val !== '' : val !== flagRaisingFilter) {
        continue;
      }
    }

    if (calcFilter !== 'all') {
      const val = row.values['計算'] || '0';
      if (val !== calcFilter) {
        continue;
      }
    }

    if (sectorFilter !== 'all' && row.sectorCode !== sectorFilter) {
      continue;
    }

    if (categoryFilter !== 'all' && row.category !== categoryFilter) {
      continue;
    }

    if (groupFilter !== 'all' && row.values['組別'] !== groupFilter) {
      continue;
    }

    if (selectionFilter === 'selected' && !selectedSet?.has(row.rowNumber)) {
      continue;
    }
    if (selectionFilter === 'unselected' && selectedSet?.has(row.rowNumber)) {
      continue;
    }

    if (externalMatchedSet !== null && !externalMatchedSet.has(row.rowNumber)) {
      continue;
    }

    if (effectiveRules.length > 0) {
      const ruleMatched =
        rulesMode === 'OR'
          ? effectiveRules.some((rule) => evaluateCustomRule(row, rule))
          : effectiveRules.every((rule) => evaluateCustomRule(row, rule));
      if (!ruleMatched) {
        continue;
      }
    }

    if (query) {
      const haystack =
        row.searchText ??
        [row.currentSectionLabel ?? '', ...Object.values(row.values)].join(' ').toLowerCase();
      if (!matchesSearchQuery(haystack, query)) {
        continue;
      }
    }

    matchedRecordRowNumbers.add(row.rowNumber);
    for (const ancRowNum of row.ancestorSectionRows ?? []) {
      matchedSectionRowNumbers.add(ancRowNum);
    }
  }

  // When sorting by a column, sort only the matched record rows for inspection
  if (sortColumn) {
    const recordResults = rows.filter(
      (r) => r.kind === 'record' && matchedRecordRowNumbers.has(r.rowNumber),
    );
    const dirMultiplier = sortDirection === 'asc' ? 1 : -1;
    return recordResults.sort((a, b) => {
      const valA = (a.values[sortColumn] ?? '').trim();
      const valB = (b.values[sortColumn] ?? '').trim();
      if (valA === '' && valB === '') {
        return a.rowNumber - b.rowNumber;
      }
      if (valA === '') return 1;
      if (valB === '') return -1;
      const cmp = zhCollator.compare(valA, valB);
      if (cmp !== 0) {
        return cmp * dirMultiplier;
      }
      return a.rowNumber - b.rowNumber;
    });
  }

  return rows.filter((row) => {
    if (row.kind === 'section') {
      if (!options.showSectionHeaders) {
        return false;
      }
      if (!hasAnyFilter) {
        return true;
      }
      if (!hasNonSearchFilter) {
        const haystack =
          row.searchText ??
          [row.sectionCode ?? '', row.sectionTitle ?? '', ...Object.values(row.values)]
            .join(' ')
            .toLowerCase();
        return matchedSectionRowNumbers.has(row.rowNumber) || matchesSearchQuery(haystack, query);
      }
      return matchedSectionRowNumbers.has(row.rowNumber);
    }

    return matchedRecordRowNumbers.has(row.rowNumber);
  });
}

export function getSectionRecordRowNumbers(
  rows: DbRowItem[],
  sectionRowNumber: number,
): number[] {
  const result: number[] = [];
  for (const row of rows) {
    if (row.kind === 'record' && row.ancestorSectionRows?.includes(sectionRowNumber)) {
      result.push(row.rowNumber);
    }
  }
  return result;
}

export function computeDbFacetCounts(rows: DbRowItem[]): {
  bySector: Map<string, number>;
  byCategory: Map<string, number>;
  byGroup: Map<string, number>;
} {
  const bySector = new Map<string, number>();
  const byCategory = new Map<string, number>();
  const byGroup = new Map<string, number>();

  for (const row of rows) {
    if (row.kind !== 'record') {
      continue;
    }
    if (row.sectorCode) {
      bySector.set(row.sectorCode, (bySector.get(row.sectorCode) ?? 0) + 1);
    }
    if (row.category) {
      byCategory.set(row.category, (byCategory.get(row.category) ?? 0) + 1);
    }
    const groupCode = row.values['組別'] ?? '';
    if (groupCode) {
      byGroup.set(groupCode, (byGroup.get(groupCode) ?? 0) + 1);
    }
  }

  return { bySector, byCategory, byGroup };
}

export interface GroupFilterOptions {
  searchQuery: string;
  sectorFilter: string;
  categoryFilter: string;
}

export function filterGroups(groups: CrmGroupItem[], options: GroupFilterOptions): CrmGroupItem[] {
  const query = options.searchQuery.trim().toLowerCase();

  return groups.filter((group) => {
    if (options.sectorFilter !== 'all' && group.sectorCode !== options.sectorFilter) {
      return false;
    }
    if (options.categoryFilter !== 'all' && group.category !== options.categoryFilter) {
      return false;
    }
    if (!query) {
      return true;
    }
    return (
      group.code.toLowerCase().includes(query) ||
      group.sectorCode.toLowerCase().includes(query) ||
      group.category.toLowerCase().includes(query) ||
      group.nameZh.toLowerCase().includes(query) ||
      group.id.toLowerCase().includes(query)
    );
  });
}
