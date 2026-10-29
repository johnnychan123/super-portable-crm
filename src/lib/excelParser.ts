import * as XLSX from 'xlsx';

export interface DbRowItem {
  rowNumber: number;
  kind: 'section' | 'record';
  sectionCode?: string;
  sectionTitle?: string;
  currentSectionLabel?: string;
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

        const groupVal = values['組別'] ?? '';
        const nameVal = values['姓名'] ?? '';
        const calcVal = values['計算'] ?? '';
        const attendVal = values['出席活動'] ?? '';
        const match = groupVal === '' && calcVal === '' && attendVal === '' ? SECTION_CODE_REGEX.exec(nameVal) : null;

        if (match) {
          const sectionCode = match[1];
          const sectionTitle = match[2]?.trim() ?? '';
          currentSectionLabel = nameVal;
          if (sectionTitle && !/^X+$/i.test(sectionTitle)) {
            sectionNameMap.set(sectionCode, sectionTitle);
          }
          dbRows.push({
            rowNumber: r + 1,
            kind: 'section',
            sectionCode,
            sectionTitle,
            currentSectionLabel,
            values,
          });
        } else {
          dbRows.push({
            rowNumber: r + 1,
            kind: 'record',
            currentSectionLabel,
            values,
          });
        }
      }
    }
  }

  // Parse 組別 sheet
  const groups: CrmGroupItem[] = [];
  const groupSheet = workbook.Sheets['組別'];
  if (groupSheet) {
    const groupMatrix = XLSX.utils.sheet_to_json<unknown[]>(groupSheet, {
      header: 1,
      defval: '',
      blankrows: false,
    });
    for (let r = 1; r < groupMatrix.length; r++) {
      const row = groupMatrix[r] ?? [];
      const id = normalizeCellText(row[0]);
      const sectorCode = normalizeCellText(row[1]);
      const category = normalizeCellText(row[2]);
      const code = normalizeCellText(row[3]);
      const rawNameZh = normalizeCellText(row[4]);
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
  const sectorSheet = workbook.Sheets['界別'];
  if (sectorSheet) {
    const sectorMatrix = XLSX.utils.sheet_to_json<unknown[]>(sectorSheet, {
      header: 1,
      defval: '',
      blankrows: false,
    });
    for (let r = 1; r < sectorMatrix.length; r++) {
      const row = sectorMatrix[r] ?? [];
      const id = normalizeCellText(row[0]);
      const code = normalizeCellText(row[1]);
      const rawNameZh = normalizeCellText(row[2]);
      const remark = normalizeCellText(row[3]);
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

export interface DbFilterOptions {
  searchQuery: string;
  attendanceFilter: 'all' | 'Y' | '不出席活動';
  showSectionHeaders: boolean;
}

export function filterDbRows(rows: DbRowItem[], options: DbFilterOptions): DbRowItem[] {
  const query = options.searchQuery.trim().toLowerCase();

  return rows.filter((row) => {
    if (row.kind === 'section') {
      if (!options.showSectionHeaders) {
        return false;
      }
      if (options.attendanceFilter !== 'all') {
        return false;
      }
      if (!query) {
        return true;
      }
      const sectionMatch =
        (row.sectionCode?.toLowerCase().includes(query) ?? false) ||
        (row.sectionTitle?.toLowerCase().includes(query) ?? false) ||
        Object.values(row.values).some((v) => v.toLowerCase().includes(query));
      return sectionMatch;
    }

    if (
      options.attendanceFilter !== 'all' &&
      row.values['出席活動'] !== options.attendanceFilter
    ) {
      return false;
    }

    if (!query) {
      return true;
    }

    return (
      (row.currentSectionLabel?.toLowerCase().includes(query) ?? false) ||
      Object.values(row.values).some((v) => v.toLowerCase().includes(query))
    );
  });
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
