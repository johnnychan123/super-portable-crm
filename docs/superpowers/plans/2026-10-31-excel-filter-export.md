# 大型名單篩選、選取、排序與原結構 Excel 匯出 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 為 Super Portable CRM 新增支援一萬多筆名單的流暢多維度篩選、階層導覽、自訂複合條件、欄位檢視排序、批次勾選與 100% 維持原檔三工作表結構（含上層分組標題列與合併儲存格）的 Excel 匯出功能。

**Architecture:** 在 `src/lib/excelParser.ts` 建立三層分組標題（界別 `MA000` -> 主組 `MA001` -> 子組 `MA001-01`）關聯索引與多維度篩選/排序引擎；在 `src/lib/excelExporter.ts` 實作依原檔列順序與命中記錄回推上層分組標題列（含 `G:H` 合併儲存格）並保留完整 `組別`/`界別` 工作表的匯出器；在 `src/components/VirtualizedDbTable.tsx` 與 `src/App.tsx` 實作固定列高虛擬捲動表格、左側界別/分類/組別導覽面板、自訂欄位複合篩選器與雙模式（Electron IPC + 瀏覽器 Blob）檔案匯出。

**Tech Stack:** TypeScript, React 19, Tailwind CSS v4, SheetJS (`xlsx`), Electron (`ipcMain` / `contextBridge`), Bun Test.

**Spec:** `docs/superpowers/specs/2026-10-31-excel-filter-export-design.md`

## Global Constraints

- `package.json` 的 `"dependencies": {}` 必須保持為空，不可新增執行期外部相依套件。
- `electron/*.ts` 嚴禁使用 `__dirname` 或 `__filename`。
- 所有新增 Electron IPC 方法必須同步定義於 `electron/main.ts`、`electron/preload.ts` 與 `src/vite-env.d.ts`。
- 匯出 Excel 時永遠維持原檔列順序（`rowNumber` 升冪）與完整 50 個 `DB` 欄位，即使介面處於部分欄位檢視或欄位排序狀態。
- 匯出 `DB` 工作表時，僅保留「有命中匯出名單」所屬的上層分組標題列（含頂層界別、母組別、子組別），並為每個匯出的分組標題列加上第 7、8 欄（`c: 6` 至 `c: 7`）合併儲存格。

## Review Focus

1. **三層分組標題列回推 (Multi-level ancestor section retention)**：當只匯出子組別（如 `MA001-02`）的一筆記錄時，其所屬子組標題列（Row 15 `MA001-02`）、母組標題列（Row 3 `MA001`）與頂層界別標題列（Row 2 `MA000`）都必須一併保留，且同屬 `MA001` 但無命中記錄的 `MA001-01`、`MA001-03` 標題列必須被排除。
2. **介面排序狀態下匯出仍維持原檔順序 (Export order invariance under active UI sort)**：當使用者在介面設定依欄位降冪排序後匯出，產出的 Excel 仍須嚴格依原檔 `rowNumber` 升冪排列並保留對應分組標題列。
3. **空白值與特殊符號篩選 (Blank & symbol filter accuracy)**：篩選「未填出席 (`blank`)」、「酒會 (`%`)」、「升旗 (`*`)」與自訂條件 `is_empty` / `is_not_empty` 時，不得誤納分組標題列或漏算空字串記錄。
4. **匯出檔案再次匯入的完整性 (Round-trip re-import fidelity)**：匯出後的 `.xlsx` Buffer 再次丟入 `parseCrmWorkbook` 時，必須能正確解析出 50 個 `dbHeaders`、保留的 `section` 與 `record` 列、558 個 `groups` 與 26 個 `sectors`。
5. **萬筆虛擬捲動邊界計算 (Virtual window boundary safety)**：當資料列為 0 筆、少於可視範圍或超過 10,000 筆且 `scrollTop` 捲至底部時，`computeVirtualWindow` 計算出的 `startIndex`、`endIndex`、`topSpacerHeight`、`bottomSpacerHeight` 不得產生越界或負數高度。

---

### Task 1: 擴充資料模型、分組階層追蹤與多維度篩選/排序引擎

**Files:**
- Modify: `src/lib/excelParser.ts`
- Test: `src/__tests__/excelParser.test.ts`

**Interfaces:**
- Produces in `src/lib/excelParser.ts`:
  - `DbRowItem`: 新增 `ancestorSectionRows: number[]`, `sectorCode?: string`, `category?: string`, `groupNameZh?: string`, `searchText?: string`
  - `ParsedCrmWorkbook`: 新增 `rawGroupRows: string[][]`, `rawSectorRows: string[][]`
  - `CustomFilterOperator = 'contains' | 'not_contains' | 'equals' | 'not_equals' | 'is_empty' | 'is_not_empty'`
  - `CustomFilterRule = { id: string; column: string; operator: CustomFilterOperator; value: string }`
  - `DbSortOption = { column: string | null; direction: 'asc' | 'desc' }`
  - `DbFilterOptions`: 支援 `searchQuery`, `attendanceFilter` (`'all' | 'Y' | '不出席活動' | 'blank'`), `receptionFilter` (`'all' | '%' | 'blank'`), `flagRaisingFilter` (`'all' | '*' | 'blank'`), `calcFilter` (`'all' | '1' | '0'`), `sectorFilter`, `categoryFilter`, `groupFilter`, `selectionFilter` (`'all' | 'selected' | 'unselected'`), `selectedRowNumbers?: ReadonlySet<number>`, `customRules?: CustomFilterRule[]`, `customRulesMode?: 'AND' | 'OR'`, `showSectionHeaders: boolean`, `sort?: DbSortOption`
  - `getSectionRecordRowNumbers(rows: DbRowItem[], sectionRowNumber: number): number[]`
  - `computeDbFacetCounts(rows: DbRowItem[]): { bySector: Map<string, number>; byCategory: Map<string, number>; byGroup: Map<string, number> }`

- [ ] **Step 1: 在 `src/__tests__/excelParser.test.ts` 撰寫分組階層追蹤、多維度篩選、自訂規則與排序測試**

```ts
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
    expect(sortedDesc[0]?.values['出席活動']).toBe('不出席活動');
    expect(sortedDesc[sortedDesc.length - 1]?.values['出席活動']).toBe('Y');
  });
```

- [ ] **Step 2: 執行測試確認新測試失敗**

Run: `bun test src/__tests__/excelParser.test.ts`
Expected: FAIL (`ancestorSectionRows` undefined)

- [ ] **Step 3: 在 `src/lib/excelParser.ts` 實作三層分組標題追蹤、擴充篩選器、排序與統計輔助函式**

實作要點：
1. 先解析 `組別` 與 `界別`（保存 `rawGroupRows` 與 `rawSectorRows`），或解析完 `DB` 回填名稱後再將每個 `record` 列關聯至 `groupMap.get(groupCode)` 以填入 `sectorCode`、`category`、`groupNameZh` 與小寫 `searchText`。
2. 在 `DB` 迴圈維護 `activeLevel0Row`（`sectionCode.endsWith('000')`）、`activeLevel1Row`（不含 `-` 且非 `000` 結尾）、`activeLevel2Row`（含 `-`，如 `MA001-01`；若 `activeLevel1Row` 的 `sectionCode` 不是此子組的前綴則清除 `activeLevel1Row`）。
3. `filterDbRows` 支援多關鍵字（空白分隔 AND、`|` 分隔 OR）、快速篩選、自訂規則（`contains`, `not_contains`, `equals`, `not_equals`, `is_empty`, `is_not_empty`）、勾選篩選、智慧保留 `matchedSectionRowNumbers`，以及 `Intl.Collator('zh-Hant', { numeric: true, sensitivity: 'base' })` 排序（空字串排在最後）。

- [ ] **Step 4: 執行測試確認通過**

Run: `bun test src/__tests__/excelParser.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/lib/excelParser.ts src/__tests__/excelParser.test.ts
git commit -m "feat(parser): add ancestor section tracking, compound filters, and sorting"
```

---

### Task 2: 實作維持原結構與上層分組標題列的 Excel 匯出引擎

**Files:**
- Create: `src/lib/excelExporter.ts`
- Create: `src/__tests__/excelExporter.test.ts`

**Interfaces:**
- Consumes from `src/lib/excelParser.ts`: `ParsedCrmWorkbook`, `DbRowItem`, `parseCrmWorkbook`
- Produces in `src/lib/excelExporter.ts`:
  - `ExportCrmOptions = { targetRecordRowNumbers: ReadonlySet<number> }`
  - `ExportPreviewSummary = { rowsToExport: DbRowItem[]; recordCount: number; sectionCount: number }`
  - `resolveExportDbRows(workbook: ParsedCrmWorkbook, options: ExportCrmOptions): ExportPreviewSummary`
  - `buildExportWorkbookBuffer(workbook: ParsedCrmWorkbook, options: ExportCrmOptions): Uint8Array`
  - `buildDefaultExportFileName(originalFileName: string, suffix?: string, now?: Date): string`

- [ ] **Step 1: 建立 `src/__tests__/excelExporter.test.ts` 驗證原結構匯出、分組標題列保留、合併儲存格與再次匯入 (Round-Trip)**

```ts
import { describe, expect, test } from 'bun:test';
import fs from 'node:fs';
import path from 'node:path';
import * as XLSX from 'xlsx';
import { parseCrmWorkbook } from '../lib/excelParser';
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

  test('builds valid .xlsx buffer preserving all 50 DB headers, G:H section merges, 組別 and 界別 sheets', () => {
    const outBuffer = buildExportWorkbookBuffer(workbook, {
      targetRecordRowNumbers: new Set([5, 16, 50]),
    });

    const rawExported = XLSX.read(outBuffer, { type: 'array' });
    expect(rawExported.SheetNames).toEqual(['DB', '組別', '界別']);
    const dbSheet = rawExported.Sheets['DB'];
    expect(dbSheet['!merges']?.length).toBe(6); // 6 section rows exported: 2, 3, 4, 15, 48, 49

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
});
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `bun test src/__tests__/excelExporter.test.ts`
Expected: FAIL (Cannot find module `'../lib/excelExporter'`)

- [ ] **Step 3: 實作 `src/lib/excelExporter.ts`**

實作要點：
1. `resolveExportDbRows`：走訪 `workbook.dbRows`，找出 `kind === 'record' && options.targetRecordRowNumbers.has(row.rowNumber)` 的列，將其 `rowNumber` 與所有 `ancestorSectionRows` 加入保留集合，最後過濾 `workbook.dbRows`（天然維持原檔 `rowNumber` 升冪順序）。
2. `buildExportWorkbookBuffer`：
   - 建立 `DB` 二維陣列：第 0 列為 `workbook.dbHeaders`，後續每列依 `workbook.dbHeaders.map(h => row.values[h] ?? '')` 填入。
   - 記錄每個 `kind === 'section'` 在輸出陣列中的 0-based 列號 `r`，設定 `dbSheet['!merges'] = sectionOutputIndices.map(r => ({ s: { c: 6, r }, e: { c: 7, r } }))`。
   - 由 `workbook.rawGroupRows` 與 `workbook.rawSectorRows` 建立 `組別` 與 `界別` 工作表並透過 `XLSX.write(wb, { type: 'array', bookType: 'xlsx' })` 回傳 `new Uint8Array(arrayBuffer)`。

- [ ] **Step 4: 執行測試確認通過**

Run: `bun test src/__tests__/excelExporter.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/lib/excelExporter.ts src/__tests__/excelExporter.test.ts
git commit -m "feat(export): add structure-preserving Excel exporter with section merges"
```

---

### Task 3: 新增 Electron 儲存檔案 IPC 與萬筆虛擬捲動資料表元件

**Files:**
- Modify: `electron/main.ts`
- Modify: `electron/preload.ts`
- Modify: `src/vite-env.d.ts`
- Create: `src/components/VirtualizedDbTable.tsx`
- Test: `src/__tests__/app.test.ts`

**Interfaces:**
- Produces in `electron/preload.ts` & `src/vite-env.d.ts`:
  - `ExcelSaveFilePayload = { defaultFileName: string; buffer: Uint8Array }`
  - `ExcelSaveFileResult = { canceled: boolean; filePath?: string }`
  - `ElectronAPI.saveExcelFile?: (payload: ExcelSaveFilePayload) => Promise<ExcelSaveFileResult>`
- Produces in `src/components/VirtualizedDbTable.tsx`:
  - `computeVirtualWindow(scrollTop: number, viewportHeight: number, rowHeight: number, totalRows: number, overscan?: number): { startIndex: number; endIndex: number; topSpacerHeight: number; bottomSpacerHeight: number }`
  - `VirtualizedDbTable` React component

- [ ] **Step 1: 在 `src/__tests__/app.test.ts` 新增虛擬捲動視窗邊界測試與 IPC 檢查**

```ts
import { computeVirtualWindow } from '../components/VirtualizedDbTable';

  test('computeVirtualWindow calculates safe start/end indices and spacer heights for 10,000+ rows', () => {
    const emptyWin = computeVirtualWindow(0, 600, 40, 0, 10);
    expect(emptyWin).toEqual({
      startIndex: 0,
      endIndex: 0,
      topSpacerHeight: 0,
      bottomSpacerHeight: 0,
    });

    const midWin = computeVirtualWindow(4000, 600, 40, 12000, 10);
    expect(midWin.startIndex).toBe(90);
    expect(midWin.endIndex).toBe(125);
    expect(midWin.topSpacerHeight).toBe(90 * 40);
    expect(midWin.bottomSpacerHeight).toBe((12000 - 125) * 40);
  });
```

- [ ] **Step 2: 執行測試確認失敗**

Run: `bun test src/__tests__/app.test.ts`
Expected: FAIL (Cannot find module `'../components/VirtualizedDbTable'`)

- [ ] **Step 3: 實作 `electron/main.ts`、`electron/preload.ts`、`src/vite-env.d.ts` 的 `excel:save-file` IPC 與 `src/components/VirtualizedDbTable.tsx`**

實作要點：
1. 在 `electron/main.ts` 註冊 `ipcMain.handle('excel:save-file', async (_event, payload: { defaultFileName: string; buffer: Uint8Array }) => ...)`，使用 `dialog.showSaveDialog` 與 `fs.writeFile` 寫入檔案，維持不使用 `__dirname`。
2. 在 `electron/preload.ts` 與 `src/vite-env.d.ts` 暴露 `saveExcelFile`。
3. 在 `src/components/VirtualizedDbTable.tsx` 實作 `computeVirtualWindow` 與 `VirtualizedDbTable`：
   - 支援每列 Checkbox 勾選、分組標題列整組全選/部分選取 Checkbox、點擊欄位標頭三態排序切換（原序 -> 升冪 -> 降冪）。
   - 使用上下 Spacer `<TableRow>` 確保 DOM 永遠只渲染可視範圍附近的列數。

- [ ] **Step 4: 執行測試與型別檢查確認通過**

Run: `bun test src/__tests__/app.test.ts && bun run typecheck`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add electron/main.ts electron/preload.ts src/vite-env.d.ts src/components/VirtualizedDbTable.tsx src/__tests__/app.test.ts
git commit -m "feat(ui): add VirtualizedDbTable and excel:save-file Electron IPC"
```

---

### Task 4: 整合左側界別/組別導覽面板、自訂複合篩選器、批次選取與匯出流程至 `src/App.tsx`

**Files:**
- Modify: `src/App.tsx`
- Modify: `src/__tests__/app.test.ts`

**Interfaces:**
- Consumes:
  - `src/lib/excelParser.ts` (`filterDbRows`, `computeDbFacetCounts`, `getSectionRecordRowNumbers`, `CustomFilterRule`, `DbSortOption`)
  - `src/lib/excelExporter.ts` (`resolveExportDbRows`, `buildExportWorkbookBuffer`, `buildDefaultExportFileName`)
  - `src/components/VirtualizedDbTable.tsx` (`VirtualizedDbTable`)

- [ ] **Step 1: 在 `src/__tests__/app.test.ts` 撰寫批次勾選邏輯與匯出整合測試**

測試驗證：
- 篩選條件組合後呼叫 `resolveExportDbRows` 針對「目前篩選結果」與「僅已勾選名單」的匯出筆數計算正確。

- [ ] **Step 2: 執行測試確認狀態**

Run: `bun test src/__tests__/app.test.ts`

- [ ] **Step 3: 更新 `src/App.tsx` 實作完整介面**

包含：
1. **頂部匯出操作區**：顯示「匯出篩選結果 (N 筆)」與「匯出已勾選 (M 筆)」按鈕，支援 Electron `window.electronAPI.saveExcelFile` 與瀏覽器 `<a download>` 雙模式匯出，並顯示成功儲存提示訊息。
2. **左側可收合導覽面板 (Facet Sidebar)**：可依「界別 (26)」、「職務分類 (14)」、「組別 (558，含搜尋框)」快速篩選並提供整組勾選 Checkbox。
3. **頂部快速篩選與自訂複合篩選器**：
   - 多關鍵字搜尋、出席狀態（全部 / Y / 不出席 / 未填）、酒會 `%`、升旗 `*`、計算 `1 / 0`、勾選狀態（全部 / 已勾選 / 未勾選）。
   - 可展開的「自訂欄位條件」面板（支援對 50 欄任意新增規則、`AND / OR` 切換、一鍵清除所有篩選）。
4. **批次選取與排序狀態列**：提供「全選目前篩選」、「取消目前篩選」、「反選目前篩選」、「清空已選」，以及目前欄位排序指示與「恢復原檔順序」按鈕。

- [ ] **Step 4: 執行完整測試、型別檢查與建置驗證**

Run: `bun run typecheck && bun test && bun run build`
Expected: All tests pass, zero TypeScript errors, both `dist/` and `dist-electron/` build cleanly.

- [ ] **Step 5: Commit**

```bash
git add src/App.tsx src/__tests__/app.test.ts docs/superpowers/specs/2026-10-31-excel-filter-export-design.md docs/superpowers/plans/2026-10-31-excel-filter-export.md
git commit -m "feat: integrate facet sidebar, compound filters, batch selection, and Excel export"
```
