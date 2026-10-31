import * as XLSX from 'xlsx';
import type { DbRowItem, ParsedCrmWorkbook } from './excelParser';

export interface ExportCrmOptions {
  targetRecordRowNumbers: ReadonlySet<number>;
  outputHeaders?: string[];
  populateExtraColumnValues?: boolean;
  extraValuesByRowNumber?: ReadonlyMap<number, Record<string, string>>;
}

export interface ExportPreviewSummary {
  rowsToExport: DbRowItem[];
  recordCount: number;
  sectionCount: number;
}

const DB_HEADER_WIDTHS: Record<string, number> = {
  酒會: 5,
  升旗: 5,
  計算: 5,
  組別: 11,
  姓名: 25,
  職位: 60,
};

const CENTERED_DB_HEADERS = new Set(['酒會', '升旗', '計算']);

const PAGE_MARGINS: XLSX.MarginInfo = {
  left: 0.7,
  right: 0.7,
  top: 0.75,
  bottom: 0.75,
  header: 0.3,
  footer: 0.3,
};

const TEST_XLSX_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" mc:Ignorable="x14ac x16r2 xr"><fonts count="8" x14ac:knownFonts="1"><font><sz val="11"/><color indexed="8"/><name val="新細明體"/><family val="2"/><scheme val="minor"/></font><font><sz val="12"/><name val="Times New Roman"/><family val="1"/></font><font><sz val="12"/><color indexed="10"/><name val="Times New Roman"/><family val="1"/></font><font><b/><sz val="12"/><name val="Times New Roman"/><family val="1"/></font><font><sz val="12"/><name val="Times New Roman"/><family val="1"/></font><font><b/><sz val="12"/><name val="Times New Roman"/><family val="1"/></font><font><sz val="12"/><name val="Times New Roman"/><family val="1"/></font><font><b/><sz val="12"/><name val="Times New Roman"/><family val="1"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="lightGray"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="15"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="3" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf><xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="3" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="一般" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext><ext xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" uri="{9260A510-F301-46a8-8635-F512D64BE5F5}"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>`;

function buildParentGroupPrefixSet(workbook: ParsedCrmWorkbook): Set<string> {
  const prefixes = new Set<string>();
  for (const group of workbook.groups) {
    const dashIdx = group.code.indexOf('-');
    if (dashIdx > 0) {
      prefixes.add(group.code.slice(0, dashIdx));
    }
  }
  for (const row of workbook.dbRows) {
    const code = row.kind === 'section' ? (row.sectionCode ?? '') : (row.values['組別'] ?? '');
    const dashIdx = code.indexOf('-');
    if (dashIdx > 0) {
      prefixes.add(code.slice(0, dashIdx));
    }
  }
  return prefixes;
}

function isUpperLevelSectionRow(row: DbRowItem, parentGroupPrefixes: ReadonlySet<string>): boolean {
  const code = row.sectionCode ?? '';
  if (code.endsWith('000')) {
    return true;
  }
  if (!code.includes('-') && parentGroupPrefixes.has(code)) {
    return true;
  }
  return false;
}

function buildColsXml(colWidths: Array<number | undefined>): string {
  const ranges: Array<{ min: number; max: number; width: number }> = [];
  for (let i = 0; i < colWidths.length; i++) {
    const w = colWidths[i];
    if (w === undefined) continue;
    const colNum = i + 1;
    const prev = ranges[ranges.length - 1];
    if (prev && prev.max === colNum - 1 && prev.width === w) {
      prev.max = colNum;
    } else {
      ranges.push({ min: colNum, max: colNum, width: w });
    }
  }
  if (ranges.length === 0) {
    return '';
  }
  const items = ranges
    .map((r) => `<col min="${r.min}" max="${r.max}" width="${r.width}" customWidth="1"/>`)
    .join('');
  return `<cols>${items}</cols>`;
}

function applyWorksheetXmlFormatting(
  sheetXml: string,
  options: {
    tabSelected?: boolean;
    colWidths: Array<number | undefined>;
    matrix: Array<Array<string | number>>;
    getCellStyleIndex: (rowIdx0: number, colIdx0: number) => number;
  },
): string {
  const tabSelectedAttr = options.tabSelected ? ' tabSelected="1"' : '';
  const sheetViewsXml = `<sheetViews><sheetView${tabSelectedAttr} workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15.75"/>`;

  let updated = sheetXml.replace(/<sheetViews>[\s\S]*?<\/sheetViews>/, sheetViewsXml);

  const colsXml = buildColsXml(options.colWidths);
  if (/<cols>[\s\S]*?<\/cols>/.test(updated)) {
    updated = updated.replace(/<cols>[\s\S]*?<\/cols>/, colsXml);
  } else if (colsXml) {
    updated = updated.replace('<sheetData>', `${colsXml}<sheetData>`);
  }

  updated = updated.replace(
    /<c r="([A-Z]+)(\d+)"([^>]*?)>([\s\S]*?)<\/c>/g,
    (_full, colLetters: string, rowNumStr: string, restAttrs: string, inner: string) => {
      const rowIdx0 = Number(rowNumStr) - 1;
      const colIdx0 = XLSX.utils.decode_col(colLetters);
      const styleIdx = options.getCellStyleIndex(rowIdx0, colIdx0);
      const rawCellVal = options.matrix[rowIdx0]?.[colIdx0];
      if (rawCellVal === '' || rawCellVal === undefined) {
        return `<c r="${colLetters}${rowNumStr}" s="${styleIdx}"/>`;
      }
      const cleanedAttrs = restAttrs.replace(/\bs="[^"]*"/g, '');
      return `<c r="${colLetters}${rowNumStr}" s="${styleIdx}"${cleanedAttrs}>${inner}</c>`;
    },
  );

  // Strip SheetJS <ignoredErrors> tag to match test.xlsx clean structure
  updated = updated.replace(/<ignoredErrors>[\s\S]*?<\/ignoredErrors>/g, '');

  return updated;
}

export function resolveExportDbRows(
  workbook: ParsedCrmWorkbook,
  options: ExportCrmOptions,
): ExportPreviewSummary {
  const includedRowNumbers = new Set<number>();
  let recordCount = 0;

  for (const row of workbook.dbRows) {
    if (row.kind === 'record' && options.targetRecordRowNumbers.has(row.rowNumber)) {
      includedRowNumbers.add(row.rowNumber);
      recordCount++;
      for (const ancRowNum of row.ancestorSectionRows ?? []) {
        includedRowNumbers.add(ancRowNum);
      }
    }
  }

  const rowsToExport = workbook.dbRows
    .filter((row) => includedRowNumbers.has(row.rowNumber))
    .sort((a, b) => a.rowNumber - b.rowNumber);

  const sectionCount = rowsToExport.filter((r) => r.kind === 'section').length;

  return {
    rowsToExport,
    recordCount,
    sectionCount,
  };
}

export function buildExportWorkbookBuffer(
  workbook: ParsedCrmWorkbook,
  options: ExportCrmOptions,
): Uint8Array {
  const { rowsToExport } = resolveExportDbRows(workbook, options);
  const outWb = XLSX.utils.book_new();

  const headersToUse =
    options.outputHeaders && options.outputHeaders.length > 0
      ? options.outputHeaders
      : workbook.dbHeaders;
  const dbHeaderSet = new Set(workbook.dbHeaders);
  const nameColIdx = headersToUse.indexOf('姓名');
  const hasAdjacentTitleCol = nameColIdx >= 0 && headersToUse[nameColIdx + 1] === '職位';
  const parentGroupPrefixes = buildParentGroupPrefixSet(workbook);

  // 1. Build DB worksheet with selected headers in order
  const dbMatrix: Array<Array<string | number>> = [headersToUse];
  const sectionMerges: XLSX.Range[] = [];

  for (let i = 0; i < rowsToExport.length; i++) {
    const row = rowsToExport[i];
    const outRowIndex = i + 1; // 0-based index in worksheet (row 0 is header)

    if (row.kind === 'section') {
      const rowCells: Array<string | number> = headersToUse.map((header) =>
        dbHeaderSet.has(header) ? (row.values[header] ?? '') : '',
      );

      if (nameColIdx >= 0) {
        if (hasAdjacentTitleCol) {
          sectionMerges.push({
            s: { c: nameColIdx, r: outRowIndex },
            e: { c: nameColIdx + 1, r: outRowIndex },
          });
        }
      } else if (rowCells.length > 0 && rowCells[0] === '') {
        rowCells[0] =
          row.values['姓名'] || row.currentSectionLabel || row.sectionCode || '';
      }

      dbMatrix.push(rowCells);
    } else {
      const extraRowVals = options.populateExtraColumnValues
        ? options.extraValuesByRowNumber?.get(row.rowNumber)
        : undefined;
      const rowCells: Array<string | number> = headersToUse.map((header) => {
        const text = dbHeaderSet.has(header)
          ? (row.values[header] ?? '')
          : (extraRowVals?.[header] ?? '');
        if (header === '計算' && /^-?\d+$/.test(text)) {
          return Number(text);
        }
        return text;
      });
      dbMatrix.push(rowCells);
    }
  }

  const dbColWidths: Array<number | undefined> = headersToUse.map(
    (header) => DB_HEADER_WIDTHS[header],
  );
  const dbSheet = XLSX.utils.aoa_to_sheet(dbMatrix);
  dbSheet['!cols'] = dbColWidths.map((w) => (w !== undefined ? { width: w } : undefined)) as XLSX.ColInfo[];
  dbSheet['!margins'] = { ...PAGE_MARGINS };
  if (sectionMerges.length > 0) {
    dbSheet['!merges'] = sectionMerges;
  }
  XLSX.utils.book_append_sheet(outWb, dbSheet, 'DB');

  // 2. Build 組別 worksheet preserving original rows and numeric ID column
  const rawGroupSource: string[][] =
    workbook.rawGroupRows && workbook.rawGroupRows.length > 0
      ? workbook.rawGroupRows
      : [
          ['ID', '界別', '分類', '簡稱', '中文'],
          ...workbook.groups.map((g) => [g.id, g.sectorCode, g.category, g.code, g.nameZh]),
        ];
  const groupMatrix: Array<Array<string | number>> = rawGroupSource.map((row, rIdx) =>
    row.map((cell, cIdx) => {
      if (rIdx > 0 && cIdx === 0 && /^-?\d+$/.test(cell)) {
        return Number(cell);
      }
      return cell;
    }),
  );
  const groupColWidths: Array<number | undefined> = [undefined, 10, 10, 10, 60, 60];
  const groupSheet = XLSX.utils.aoa_to_sheet(groupMatrix);
  groupSheet['!cols'] = groupColWidths.map((w) =>
    w !== undefined ? { width: w } : undefined,
  ) as XLSX.ColInfo[];
  groupSheet['!margins'] = { ...PAGE_MARGINS };
  XLSX.utils.book_append_sheet(outWb, groupSheet, '組別');

  // 3. Build 界別 worksheet preserving original rows and numeric ID column
  const rawSectorSource: string[][] =
    workbook.rawSectorRows && workbook.rawSectorRows.length > 0
      ? workbook.rawSectorRows
      : [
          ['ID', '簡稱', '中文', '備註'],
          ...workbook.sectors.map((s) => [s.id, s.code, s.nameZh, s.remark]),
        ];
  const sectorMatrix: Array<Array<string | number>> = rawSectorSource.map((row, rIdx) =>
    row.map((cell, cIdx) => {
      if (rIdx > 0 && cIdx === 0 && /^-?\d+$/.test(cell)) {
        return Number(cell);
      }
      return cell;
    }),
  );
  const sectorColWidths: Array<number | undefined> = [undefined, 10, 25];
  const sectorSheet = XLSX.utils.aoa_to_sheet(sectorMatrix);
  sectorSheet['!cols'] = sectorColWidths.map((w) =>
    w !== undefined ? { width: w } : undefined,
  ) as XLSX.ColInfo[];
  sectorSheet['!margins'] = { ...PAGE_MARGINS };
  XLSX.utils.book_append_sheet(outWb, sectorSheet, '界別');

  const rawArrayOut = XLSX.write(outWb, {
    type: 'array',
    bookType: 'xlsx',
    bookSST: true,
  }) as ArrayBuffer;

  // 4. Inject test.xlsx OpenXML styles, column widths, freeze panes, and cell style indices
  const CFB = (XLSX as unknown as { CFB: typeof XLSX.CFB }).CFB;
  const cfb = CFB.read(new Uint8Array(rawArrayOut), { type: 'array' });
  const encoder = new TextEncoder();
  const decoder = new TextDecoder();

  const stylesEntry = CFB.find(cfb, '/xl/styles.xml');
  if (stylesEntry) {
    stylesEntry.content = encoder.encode(TEST_XLSX_STYLES_XML);
  }

  const workbookEntry = CFB.find(cfb, '/xl/workbook.xml');
  if (workbookEntry) {
    const wbXml = decoder.decode(new Uint8Array(workbookEntry.content as ArrayBuffer));
    const definedNamesXml =
      '<definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">DB!$1:$1</definedName><definedName name="_xlnm.Print_Titles" localSheetId="2">界別!$1:$1</definedName><definedName name="_xlnm.Print_Titles" localSheetId="1">組別!$1:$1</definedName></definedNames>';
    if (!wbXml.includes('<definedNames>')) {
      workbookEntry.content = encoder.encode(
        wbXml.replace('</sheets>', `</sheets>${definedNamesXml}`),
      );
    }
  }

  const sheet1Entry = CFB.find(cfb, '/xl/worksheets/sheet1.xml');
  if (sheet1Entry) {
    const sheet1Xml = decoder.decode(new Uint8Array(sheet1Entry.content as ArrayBuffer));
    const formattedSheet1 = applyWorksheetXmlFormatting(sheet1Xml, {
      tabSelected: true,
      colWidths: dbColWidths,
      matrix: dbMatrix,
      getCellStyleIndex: (rowIdx0, colIdx0) => {
        if (rowIdx0 === 0) {
          return 3; // DB Header: Times New Roman 12 Bold
        }
        const exportedRow = rowsToExport[rowIdx0 - 1];
        if (!exportedRow) {
          return 1;
        }
        if (exportedRow.kind === 'section') {
          const isUpper = isUpperLevelSectionRow(exportedRow, parentGroupPrefixes);
          if (colIdx0 === nameColIdx) {
            return isUpper ? 12 : 14;
          }
          if (hasAdjacentTitleCol && colIdx0 === nameColIdx + 1) {
            return 13;
          }
          return isUpper ? 4 : 5;
        }
        const isNotAttending = (exportedRow.values['出席活動'] ?? '') === '不出席活動';
        const headerName = headersToUse[colIdx0] ?? '';
        if (CENTERED_DB_HEADERS.has(headerName)) {
          return isNotAttending ? 7 : 6;
        }
        return isNotAttending ? 2 : 1;
      },
    });
    sheet1Entry.content = encoder.encode(formattedSheet1);
  }

  const sheet2Entry = CFB.find(cfb, '/xl/worksheets/sheet2.xml');
  if (sheet2Entry) {
    const sheet2Xml = decoder.decode(new Uint8Array(sheet2Entry.content as ArrayBuffer));
    const formattedSheet2 = applyWorksheetXmlFormatting(sheet2Xml, {
      colWidths: groupColWidths,
      matrix: groupMatrix,
      getCellStyleIndex: (rowIdx0, colIdx0) => {
        if (rowIdx0 === 0) {
          return colIdx0 === 4 ? 3 : 9; // 組別 Header: E1=3, others=9 (Times New Roman 12 Bold)
        }
        return 8; // 組別 Data: Times New Roman 12
      },
    });
    sheet2Entry.content = encoder.encode(formattedSheet2);
  }

  const sheet3Entry = CFB.find(cfb, '/xl/worksheets/sheet3.xml');
  if (sheet3Entry) {
    const sheet3Xml = decoder.decode(new Uint8Array(sheet3Entry.content as ArrayBuffer));
    const formattedSheet3 = applyWorksheetXmlFormatting(sheet3Xml, {
      colWidths: sectorColWidths,
      matrix: sectorMatrix,
      getCellStyleIndex: (rowIdx0) => {
        if (rowIdx0 === 0) {
          return 11; // 界別 Header: Times New Roman 12 Bold
        }
        return 10; // 界別 Data: Times New Roman 12
      },
    });
    sheet3Entry.content = encoder.encode(formattedSheet3);
  }

  const finalZip = CFB.write(cfb, { type: 'array', fileType: 'zip' }) as ArrayBuffer;
  return new Uint8Array(finalZip);
}

export function buildDefaultExportFileName(
  originalFileName: string,
  suffix = '篩選匯出',
  now: Date = new Date(),
): string {
  const baseName = originalFileName.replace(/\.(xlsx|xls)$/i, '') || 'crm_export';
  const yyyy = String(now.getFullYear());
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  const hh = String(now.getHours()).padStart(2, '0');
  const min = String(now.getMinutes()).padStart(2, '0');
  return `${baseName}_${suffix}_${yyyy}${mm}${dd}_${hh}${min}.xlsx`;
}
