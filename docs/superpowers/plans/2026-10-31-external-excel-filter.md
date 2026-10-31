# 外部匯出 Excel 篩選匯入、多次比對選取與彈性欄位匯出 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 為 Super Portable CRM 新增「多次匯入過去匯出的 Excel 作為篩選與選取條件」及「按需自訂/沿用匯出欄位（含增刪欄位與活動專屬欄位）」功能。

**Architecture:** 新增 `src/lib/externalFilterMatcher.ts` 處理任意欄位結構的外部 Excel 解析、分組標題列剔除、自動/手動共同識別欄比對與多檔案聯集/交集計算；擴充 `src/lib/excelParser.ts` 的 `filterDbRows` 支援外部篩選檔命中列過濾；擴充 `src/lib/excelExporter.ts` 支援自訂匯出欄位順序、活動專屬額外欄位（預設留空、可選帶入舊檔資料）與分組標題列彈性對齊；並在 `electron/` 與 `src/App.tsx` 實作多檔匯入 IPC 與「外部篩選 Excel / 匯出欄位設定」雙抽屜面板。

**Tech Stack:** TypeScript, React 19, Tailwind CSS v4, SheetJS (`xlsx`), Electron (`ipcMain` / `contextBridge`), Bun Test.

**Spec:** `docs/superpowers/specs/2026-10-31-external-excel-filter-design.md`

## Global Constraints

- `package.json` 的 `"dependencies": {}` 必須保持為空，不可新增執行期外部相依套件。
- `electron/*.ts` 嚴禁使用 `__dirname` 或 `__filename`。
- 所有新增 Electron IPC 方法必須同步定義於 `electron/main.ts`、`electron/preload.ts` 與 `src/vite-env.d.ts`。
- 匯入外部篩選 Excel 時**預設不自動改動 `selectedRowNumbers`**，由使用者點擊「勾選此檔命中名單」或批次按鈕才變更勾選狀態。
- 當匯出包含非主 DB 原有的「活動專屬額外欄位」時，**預設只輸出空白欄位標題**（`populateExtraColumnValues = false`），僅在使用者開啟「帶入舊檔該欄資料」時填入對應值。

## Review Focus

1. **外部篩選檔內含分組標題列時的排除 (Ignoring section header rows in external Excel)**：過去由本系統匯出的 Excel 會包含 `MA000 公共行政部門`、`MZ001 基本法委員會` 等分組標題列；匯入作為篩選檔時必須自動略過這些分組標題列，不得將其誤算為「未命中人員記錄」。
2. **空白比對值防碰撞 (Blank key non-collision)**：當主 DB 或外部篩選檔在比對欄位（如 `PID` 或 `姓名`）為空字串時，絕對不可因為雙方皆為 `""` 而誤判為同一人並命中。
3. **自訂匯出欄位移除「姓名」或「職位」時的分組標題列回退 (Section header fallback when `姓名`/`職位` columns are omitted)**：當匯出欄位包含連續的 `姓名` 與 `職位` 時維持合併儲存格；若移除 `職位` 則不合併；若連 `姓名` 都移除，分組標題文字須自動寫入第 1 個匯出欄位，不可產生全空白的分組標題列。
4. **活動專屬額外欄位的預設留空與帶入舊值切換 (Extra column blank-by-default vs. value carryover)**：當沿用篩選檔欄位匯出且含額外欄位（如 `桌號`、`車牌`）時，預設匯出的資料列在該欄必須為空字串 `""`；開啟 `populateExtraColumnValues: true` 時則精準填入該筆 `rowNumber` 對應的外部檔欄位值，且分組標題列在額外欄位永遠保持空白。
5. **多次匯入篩選檔的聯集/交集與個別移除 (Multi-file union/intersection and removal consistency)**：匯入多個篩選 Excel 後，切換比對欄位、切換「任一命中（聯集）/ 全部命中（交集）/ 指定單一檔案」或移除某個篩選檔時，篩選結果與額外欄位映射必須即時重算且不殘留已移除檔案的狀態。

---

### Task 1: 外部篩選 Excel 解析、智慧欄位比對與多檔篩選引擎

**Files:**
- Create: `src/lib/externalFilterMatcher.ts`
- Modify: `src/lib/excelParser.ts`
- Create: `src/__tests__/externalFilterMatcher.test.ts`

**Interfaces:**
- Produces in `src/lib/externalFilterMatcher.ts`:
  - `ParsedExternalFilterSheet = { fileName: string; sheetName: string; headers: string[]; rows: Record<string, string>[]; sharedHeaders: string[]; extraHeaders: string[] }`
  - `ExternalMatchKeyOption = { key: string; label: string; columns: string[] }`
  - `ExternalFilterMatchResult = { matchKey: string; resolvedKeyColumns: string[]; matchedDbRowNumbers: Set<number>; matchedExternalRowCount: number; unmatchedExternalRowCount: number; totalExternalRowCount: number; extraValuesByDbRow: Map<number, Record<string, string>> }`
  - `ImportedFilterFile = { id: string; fileName: string; sheet: ParsedExternalFilterSheet; availableMatchKeys: ExternalMatchKeyOption[]; selectedMatchKey: string; matchResult: ExternalFilterMatchResult }`
  - `ExternalFilterViewMode = 'all' | 'union' | 'intersection' | string` (where specific `string` is an `ImportedFilterFile['id']`)
  - `parseExternalFilterWorkbook(data: Uint8Array | ArrayBuffer, fileName: string, dbHeaders: string[]): ParsedExternalFilterSheet`
  - `getAvailableMatchKeys(sheet: ParsedExternalFilterSheet, dbRows: DbRowItem[]): ExternalMatchKeyOption[]`
  - `matchExternalFilterToDb(sheet: ParsedExternalFilterSheet, dbRows: DbRowItem[], matchKey?: string): ExternalFilterMatchResult`
  - `computeExternalFilterRowSet(files: ImportedFilterFile[], mode: ExternalFilterViewMode): Set<number> | null`
  - `mergeExtraValuesFromFiles(files: ImportedFilterFile[]): Map<number, Record<string, string>>`
- Modifies in `src/lib/excelParser.ts`:
  - `DbFilterOptions`: 新增 `externalMatchedRowNumbers?: ReadonlySet<number> | null`

- [ ] **Step 1: 建立 `src/__tests__/externalFilterMatcher.test.ts` 撰寫外部篩選檔解析、分組標題列排除、空白防碰撞、自動/手動比對鍵與多檔聯集/交集測試**

```ts
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
      values: { 組別: '', 計算: '', 出席活動: '', 姓名: 'MA000 公共行政部門', 職位: '', PID: '', 'Pf.ID': '' },
    },
    {
      rowNumber: 3,
      kind: 'record',
      ancestorSectionRows: [2],
      values: { 組別: 'MA001', 計算: '1', 出席活動: 'Y', 姓名: '陳大文', 職位: '處長', PID: 'P001', 'Pf.ID': 'PF-101' },
    },
    {
      rowNumber: 4,
      kind: 'record',
      ancestorSectionRows: [2],
      values: { 組別: 'MA001', 計算: '1', 出席活動: 'Y', 姓名: '李小明', 職位: '主任', PID: 'P002', 'Pf.ID': '' },
    },
    {
      rowNumber: 5,
      kind: 'record',
      ancestorSectionRows: [2],
      values: { 組別: 'MA002', 計算: '0', 出席活動: '不出席活動', 姓名: '張三', 職位: '顧問', PID: '', 'Pf.ID': '' },
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
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `bun test src/__tests__/externalFilterMatcher.test.ts`
Expected: FAIL (`Cannot find module '../lib/externalFilterMatcher'`)

- [ ] **Step 3: 實作 `src/lib/externalFilterMatcher.ts` 並更新 `src/lib/excelParser.ts` 的 `filterDbRows`**

實作要點：
1. `parseExternalFilterWorkbook` 讀取名為 `'DB'` 的工作表或第一張工作表；對每個儲存格正規化 `_x000D_` 與空白；若某列在 `組別`、`計算`、`出席活動` 皆為空且 `姓名`（或第一個非空文字欄）符合 `/^(M[A-Z]\d{3}(?:-\d{2})?)\s*(.*)$/`，視為分組標題列略過。
2. `getAvailableMatchKeys` 產生 `'auto'`（自動偵測）、若同時含 `PID` 與 `組別` 則加入 `'PID+組別'`、若同時含 `組別` 與 `姓名` 則加入 `'組別+姓名'`，接著列出所有 `sharedHeaders` 單一欄位選項。
3. `matchExternalFilterToDb` 在 `'auto'` 時依優先序 `[['Pf.ID'], ['PID', '組別'], ['PID'], ['組別', '姓名'], ['姓名'], ['中文名(沒有稱謂)']]` 挑出第一個「在外部檔與主 DB 皆有至少一筆全部鍵欄非空」的欄位組合；比對時所有鍵欄都必須非空（`val !== ''`），大小寫不敏感比對，並收集 `extraHeaders` 的非空值至 `extraValuesByDbRow`。
4. 在 `src/lib/excelParser.ts` 的 `DbFilterOptions` 加入 `externalMatchedRowNumbers?: ReadonlySet<number> | null`；當其為非 null/undefined 的 `Set` 時，將 `hasNonSearchFilter` 視為 `true`，且僅保留 `externalMatchedRowNumbers.has(row.rowNumber)` 的 `record` 列與其 `ancestorSectionRows`。

- [ ] **Step 4: 執行測試確認通過**

Run: `bun test src/__tests__/externalFilterMatcher.test.ts src/__tests__/excelParser.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/lib/externalFilterMatcher.ts src/lib/excelParser.ts src/__tests__/externalFilterMatcher.test.ts
git commit -m "feat(matcher): add external Excel filter parser and multi-file matcher"
```

---

### Task 2: 彈性匯出欄位與活動專屬額外欄位匯出引擎

**Files:**
- Modify: `src/lib/excelExporter.ts`
- Modify: `src/__tests__/excelExporter.test.ts`

**Interfaces:**
- Modifies in `src/lib/excelExporter.ts`:
  - `ExportCrmOptions`:
    - `targetRecordRowNumbers: ReadonlySet<number>`
    - `outputHeaders?: string[]`
    - `populateExtraColumnValues?: boolean`
    - `extraValuesByRowNumber?: ReadonlyMap<number, Record<string, string>>`

- [ ] **Step 1: 在 `src/__tests__/excelExporter.test.ts` 新增自訂匯出欄位、額外活動欄位（預設留空 vs 帶入舊值）與分組標題列回退測試**

```ts
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

  test('falls back section header label to first column when 姓名 is omitted from outputHeaders', () => {
    const noNameHeaders = ['組別', '出席活動', '桌號'];
    const buf = buildExportWorkbookBuffer(workbook, {
      targetRecordRowNumbers: new Set([5]),
      outputHeaders: noNameHeaders,
    });
    const wbOut = XLSX.read(buf, { type: 'array' });
    const matrix = XLSX.utils.sheet_to_json<string[]>(wbOut.Sheets['DB'], {
      header: 1,
      defval: '',
    });
    expect(matrix[0]).toEqual(noNameHeaders);
    expect(matrix[1]?.[0]).toContain('MA000');
  });
```

- [ ] **Step 2: 執行測試確認新測試失敗**

Run: `bun test src/__tests__/excelExporter.test.ts`
Expected: FAIL (`rowsBlank[0]` has 50 headers instead of `customHeaders`)

- [ ] **Step 3: 更新 `src/lib/excelExporter.ts` 的 `buildExportWorkbookBuffer` 支援 `outputHeaders`、`populateExtraColumnValues` 與 `extraValuesByRowNumber`**

實作要點：
1. `const headersToUse = options.outputHeaders && options.outputHeaders.length > 0 ? options.outputHeaders : workbook.dbHeaders;`
2. 建立 `dbHeaderSet = new Set(workbook.dbHeaders)` 判斷每個輸出欄是否為主 DB 原有欄位。
3. 針對 `row.kind === 'section'`：
   - 若 `headersToUse` 包含 `'姓名'`（索引為 `nameIdx`），按原欄位輸出；且若 `headersToUse[nameIdx + 1] === '職位'`，加入 `{ s: { c: nameIdx, r: outRowIndex }, e: { c: nameIdx + 1, r: outRowIndex } }` 合併儲存格。
   - 若 `headersToUse` 不包含 `'姓名'`，將 `row.values['姓名'] || row.currentSectionLabel || row.sectionCode || ''` 填入 `rowCells[0]`，確保分組標題列不會變成空白列。
   - 分組標題列在非 DB 額外欄位永遠輸出 `''`。
4. 針對 `row.kind === 'record'`：
   - 若欄位屬於 `dbHeaderSet`，輸出 `row.values[header] ?? ''`。
   - 若欄位為額外欄位：當 `options.populateExtraColumnValues === true` 時輸出 `options.extraValuesByRowNumber?.get(row.rowNumber)?.[header] ?? ''`，否則輸出 `''`。

- [ ] **Step 4: 執行測試確認通過**

Run: `bun test src/__tests__/excelExporter.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/lib/excelExporter.ts src/__tests__/excelExporter.test.ts
git commit -m "feat(exporter): support custom output columns and extra event column carryover"
```

---

### Task 3: Electron 多檔匯入 IPC 與「外部篩選 Excel & 匯出欄位設定」UI 整合

**Files:**
- Modify: `electron/main.ts`
- Modify: `electron/preload.ts`
- Modify: `src/vite-env.d.ts`
- Modify: `src/App.tsx`
- Modify: `src/__tests__/app.test.ts`

**Interfaces:**
- Consumes from `src/lib/externalFilterMatcher.ts` & `src/lib/excelExporter.ts`
- Produces in `electron/main.ts`, `electron/preload.ts`, `src/vite-env.d.ts`:
  - `ExcelOpenedFileItem = { fileName: string; filePath?: string; buffer: Uint8Array }`
  - `ExcelOpenMultipleFilesResult = { canceled: boolean; files?: ExcelOpenedFileItem[] }`
  - `openMultipleExcelFiles?: () => Promise<ExcelOpenMultipleFilesResult>` (`ipcMain.handle('excel:open-multiple-files', ...)`)
- Produces in `src/App.tsx`:
  - `resolveActiveOutputHeaders(dbHeaders: string[], mode: 'db_all' | 'from_file' | 'custom', templateFileHeaders: string[], customSelectedDbHeaders: string[], customExtraHeaders: string[]): string[]`

- [ ] **Step 1: 在 `src/__tests__/app.test.ts` 新增 `resolveActiveOutputHeaders` 與 `excel:open-multiple-files` IPC 檢查測試**

```ts
  test('resolveActiveOutputHeaders supports db_all, from_file, and custom column modes', () => {
    const dbHeaders = ['組別', '姓名', '職位', 'PID'];
    expect(resolveActiveOutputHeaders(dbHeaders, 'db_all', [], [], [])).toEqual(dbHeaders);
    expect(
      resolveActiveOutputHeaders(dbHeaders, 'from_file', ['組別', '姓名', '桌號'], [], []),
    ).toEqual(['組別', '姓名', '桌號']);
    expect(
      resolveActiveOutputHeaders(
        dbHeaders,
        'custom',
        [],
        ['組別', '姓名'],
        ['桌號', '車牌', '姓名'], // duplicate '姓名' should be deduplicated
      ),
    ).toEqual(['組別', '姓名', '桌號', '車牌']);
  });
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `bun test src/__tests__/app.test.ts`
Expected: FAIL (`resolveActiveOutputHeaders` not exported from `../App`)

- [ ] **Step 3: 實作 `electron/` 多檔開啟 IPC 與 `src/App.tsx` 的「匯入篩選 Excel」及「匯出欄位設定」面板**

實作要點：
1. 在 `electron/main.ts` 註冊 `ipcMain.handle('excel:open-multiple-files', ...)`（使用 `properties: ['openFile', 'multiSelections']`），並在 `electron/preload.ts` 與 `src/vite-env.d.ts` 同步加上 `openMultipleExcelFiles`。
2. 在 `src/App.tsx` 匯出純函式 `resolveActiveOutputHeaders`，並新增狀態：
   - `importedFilterFiles: ImportedFilterFile[]`
   - `externalFilterMode: ExternalFilterViewMode` (`'all' | 'union' | 'intersection' | fileId`)
   - `showExternalFilterPanel: boolean`
   - `showExportColumnPanel: boolean`
   - `exportColumnMode: 'db_all' | 'from_file' | 'custom'`
   - `exportTemplateFileId: string`
   - `customExportDbHeaders: string[]`
   - `customExtraHeaders: string[]`
   - `newExtraHeaderInput: string`
   - `populateExtraColumnValues: boolean`（預設 `false`）
3. 在 `src/App.tsx` 提供：
   - 頂部與工具列「匯入篩選 Excel」按鈕（支援多次點擊匯入、累加至 `importedFilterFiles`，匯入時**不自動勾選**）。
   - 「外部篩選 Excel」抽屜面板：列出所有已匯入篩選檔卡片（顯示命中筆數、共同欄位、額外欄位、比對欄位下拉切換、「勾選此檔命中 (N)」、「取消勾選此檔」、「僅檢視此檔」、「套用此檔欄位結構」、移除按鈕），以及多檔「勾選全部篩選檔命中 / 取消全部」與篩選模式切換（全部顯示 / 任一命中聯集 / 全部命中交集 / 個別檔案）。
   - 「匯出欄位設定」抽屜面板：可切換「主 DB 完整 50 欄」、「沿用已匯入篩選檔欄位」、「自訂勾選 DB 欄位 + 新增活動欄位」，並提供「帶入舊篩選檔中對應人員的額外欄位資料」勾選開關（預設不勾選）。
   - 匯出按鈕與預覽摘要同步顯示目前選用的匯出欄位數（例如 `50 欄` 或 `自訂 M 欄`）。

- [ ] **Step 4: 執行型別檢查、單元測試與建置確認全部通過**

Run: `bun run typecheck && bun test && bun run build`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add electron/main.ts electron/preload.ts src/vite-env.d.ts src/App.tsx src/__tests__/app.test.ts docs/superpowers/specs/2026-10-31-external-excel-filter-design.md docs/superpowers/plans/2026-10-31-external-excel-filter.md
git commit -m "feat: add multi-file external Excel filter import and custom column export"
```
