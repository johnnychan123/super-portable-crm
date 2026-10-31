# Super Portable CRM — 大型名單篩選、選取、排序與原結構 Excel 匯出設計文件

- **日期**：2026-10-31
- **狀態**：待審核 (Draft)

---

## 1. 背景與目標 (Intent & Why)

Super Portable CRM 的核心工作流是：
1. **匯入**包含 `DB`（主名單，50 個欄位與內嵌分組標題列）、`組別`（558 個組別與職務分類）、`界別`（26 個頂層界別 `MA000`–`MZ000`）三張工作表的 Excel 活頁簿（`.xlsx` / `.xls`）。
2. 在應用程式中針對**一萬多筆（10,000+ rows）**名單進行高速**搜尋、多維度篩選、欄位排序檢視與勾選/批次挑選**。
3. **匯出**篩選後或勾選後的結果為新的 `.xlsx` Excel 檔案，且**完全維持原本的完整欄位（50 欄）與三工作表結構（`DB`、`組別`、`界別`）**，以便後續作業或再次匯入本系統使用。

---

## 2. 目標與非目標 (Goals, Non-Goals & Anti-Goals)

### 目標 (Goals)
- **萬筆資料流暢操作**：在 10,000+ 筆資料列與 15～50 欄位檢視下，搜尋、篩選、排序、捲動與批次勾選保持即時反應（無 DOM 卡頓）。
- **直覺的多維度篩選與導覽**：
  - 左側可收合「界別 / 職務分類 / 組別」導覽面板，顯示各分類即時命中筆數，支援點擊過濾與一鍵勾選整組。
  - 頂部快速篩選（關鍵字搜尋、出席狀態 `Y / 不出席活動 / 未填`、酒會 `%`、升旗 `*`、計算 `1 / 0`、勾選狀態 `全部 / 已勾選 / 未勾選`）。
  - 自訂多欄位複合條件篩選器（支援針對 50 個任意欄位設定 `包含`、`不包含`、`等於`、`不等於`、`為空`、`不為空`，並支援 `AND / OR` 邏輯組合）。
- **靈活的勾選與匯出模式**：
  - 支援單筆勾選、分組標題列整組勾選、「全選目前篩選結果」、「取消全選目前篩選結果」、「反選目前篩選結果」。
  - 匯出時可選擇「匯出目前篩選結果」或「僅匯出已勾選名單」。
- **100% 結構相容的 Excel 匯出**：
  - 匯出的 `.xlsx` 保留 `DB`（完整 50 欄）、`組別`、`界別` 三張工作表。
  - **順序保證**：匯出時永遠維持原檔列順序（`rowNumber` 升冪），介面上的欄位排序僅供檢視。
  - **智慧分組標題列保留**：匯出 `DB` 工作表時，僅保留「有命中匯出名單」所屬的上層分組標題列（例如頂層界別 `MA000`、母組別 `MA001`、子組別 `MA001-01`），並同步產生分組標題列的 `G:H`（`姓名:職位`）合併儲存格（`!merges`），剔除沒有任何命中成員的空組別標題列。
  - **匯出可重入性 (Round-Trip Compatibility)**：匯出後的 Excel 檔案可再次匯入本系統且結構、分組標題、組別與界別完全一致。

### 非目標 (Non-Goals)
- 不做線上多人協同編輯或外部資料庫連線（維持離線單機 Portable 架構）。
- 不修改或刪除 `組別` 與 `界別` 定義表的內容（匯出時完整保留原始參考表）。

### 反目標 (Anti-Goals)
- **不可因渲染 10,000+ 列 DOM 導致介面凍結**：嚴禁將一萬多筆 `<TableRow>` 同時掛載至 DOM。
- **不可在匯出時破壞原始欄位或打亂組別結構**：即使使用者目前切換在「核心 15 欄」或啟用了「依姓名排序檢視」，匯出的 Excel 仍須輸出全部 50 欄且維持原檔組別分段與列順序。

---

## 3. 技術與打包限制 (Constraints)

- **相依套件打包規則**：`package.json` 的 `"dependencies": {}` 必須保持為空，所有套件皆在建置期由 Vite (`dist/`) 與 `bun build` (`dist-electron/`) 打包，不新增原生模組。
- **Electron 路徑規範**：`electron/*.ts` 嚴禁使用 `__dirname` 或 `__filename`，所有新增 IPC 必須同時定義於 `electron/main.ts`、`electron/preload.ts` 及 `src/vite-env.d.ts`。
- **雙環境相容 (Electron + Web Fallback)**：匯入與匯出皆須同時支援 Electron 原生檔案對話框 (`dialog.showOpenDialog` / `dialog.showSaveDialog`) 與瀏覽器模式 (`<input type="file">` / `Blob` 下載)。

---

## 4. 評估方案與選擇 (Approaches)

1. **方案 A（已選擇）：左側界別/組別導覽面板 + 頂部快速與自訂複合篩選列 + 虛擬捲動資料表 (Split-Pane Facet Sidebar + Compound Filter Bar + Virtualized Table)**
   - **做法**：左側提供界別/職務分類/組別階層導覽與批次勾選；頂部提供快速篩選與自訂多欄位條件建構器；主表格採用輕量列虛擬化（Row Virtualization），僅渲染可視範圍與緩衝區列數。
   - **優點**：不會因分頁切斷同組別名單，一萬多筆捲動、搜尋、排序與跨組勾選皆流暢直覺。
2. **方案 B：頂部下拉篩選面板 + 分頁式資料表 (Paginated Table)**
   - **捨棄原因**：分頁會把同一組別切在不同頁，不利於上下檢視分組標題列與連續勾選。
3. **方案 C：欄位標頭內嵌漏斗篩選 + 虛擬捲動 (Header Popover Filters)**
   - **捨棄原因**：`DB` 有 50 個欄位，條件分散在標頭內不易總覽，且難以呈現跨表（`界別` / `組別` 職務分類）的階層統計。

---

## 5. 詳細設計規格 (Detailed Specification)

### 5.1 資料模型擴充與組別階層關聯 (`src/lib/excelParser.ts`)

匯入解析 `parseCrmWorkbook` 時，為每筆 `DbRowItem` 預先建立索引與階層關聯，確保萬筆資料篩選在數毫秒內完成：

1. **分組標題階層追蹤 (`ancestorSectionRows: number[]`)**：
   - `DB` 工作表中的 `section` 列分為三種層級：
     - **Level 0（界別標題列）**：`sectionCode` 以 `000` 結尾（如 `MA000`、`MB000`、`MZ000`）。
     - **Level 1（大組/主組標題列）**：`sectionCode` 不含 `-` 且不以 `000` 結尾（如 `MA001`、`MA002`、`MZ001`）。
     - **Level 2（子組標題列）**：`sectionCode` 含 `-` 子編號（如 `MA001-01`、`MA001-02`）。
   - 解析 `DB` 列時維護目前作用中的 `[activeLevel0Row, activeLevel1Row, activeLevel2Row]`：
     - 遇到 Level 0 `section` 列：更新 `activeLevel0Row = rowNumber`，清空 `activeLevel1Row`、`activeLevel2Row`。
     - 遇到 Level 1 `section` 列：更新 `activeLevel1Row = rowNumber`，清空 `activeLevel2Row`（若該組直接下轄成員如 `MA002`、`MZ001`，此列即為成員的直接組別標題列；若後續跟著 `MA001-01`，此列即為母組標題列）。
     - 遇到 Level 2 `section` 列：更新 `activeLevel2Row = rowNumber`（若其前綴與 `activeLevel1Row` 不同則清空不匹配的 `activeLevel1Row`）。
   - 每筆 `record` 列記錄：
     - `ancestorSectionRows: number[]`：所有屬於該筆記錄的上層 `section` 列號（例如 `MA001-01` 的成員會記錄 `[2 (MA000), 3 (MA001), 4 (MA001-01)]`；`MZ001` 的成員會記錄 `[48 (MZ000), 49 (MZ001)]`）。
     - `sectorCode: string`：優先從 `組別` 對照表查找該記錄 `values['組別']` 對應的 `sectorCode`；若無，則取自 `values['組別']`（或所屬 `sectionCode`）的前兩碼加 `000`（如 `MA001-01` -> `MA000`）。
     - `category: string`：從 `組別` 對照表查找 `values['組別']` 對應的職務分類（如 `現任官員`、`委員會`）。
     - `groupNameZh: string`：從 `組別` 對照表（含回填名稱）查找對應中文名稱。
     - `searchText: string`：預先將所有欄位值、組別代碼、組別中文名、界別代碼、職務分類轉為小寫拼接字串，供萬筆搜尋快速比對。

2. **原始工作表資料保留**：
   - 在 `ParsedCrmWorkbook` 中保存 `rawGroupRows: unknown[][]` 與 `rawSectorRows: unknown[][]`（包含標題列與完整原始儲存格內容），使匯出時能 100% 還原 `組別` 與 `界別` 工作表，不因 UI 檢視或名稱回填而遺失原始欄位。

### 5.2 搜尋、多維度篩選與檢視排序邏輯 (`src/lib/excelParser.ts`)

擴充 `DbFilterOptions` 與處理函式：

```ts
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
  column: string | null; // null 代表維持原檔順序
  direction: 'asc' | 'desc';
}

export interface DbFilterOptions {
  searchQuery: string;
  attendanceFilter: 'all' | 'Y' | '不出席活動' | 'blank';
  receptionFilter?: 'all' | '%' | 'blank';
  flagRaisingFilter?: 'all' | '*' | 'blank';
  calcFilter?: 'all' | '1' | '0';
  sectorFilter?: string; // 'all' 或特定界別代碼如 'MA000'
  categoryFilter?: string; // 'all' 或特定職務分類如 '現任官員'
  groupFilter?: string; // 'all' 或特定組別代碼如 'MZ001'
  selectionFilter?: 'all' | 'selected' | 'unselected';
  selectedRowNumbers?: ReadonlySet<number>;
  customRules?: CustomFilterRule[];
  customRulesMode?: 'AND' | 'OR';
  showSectionHeaders: boolean;
  sort?: DbSortOption;
}
```

- **篩選與分組標題列顯示規則**：
  1. 先針對所有 `kind === 'record'` 的資料列評估是否符合所有啟用的篩選條件（搜尋字串支援空白分隔多關鍵字 AND 比對，以及 `|` 分隔 OR 比對）。
  2. 收集所有命中 `record` 列的 `ancestorSectionRows` 聯集（`matchedSectionRowNumbers`）。
  3. 當 `showSectionHeaders === true` 且**未啟用欄位排序 (`!sort?.column`)** 時：
     - 若目前有任何篩選條件啟用（搜尋、出席、界別、分類、組別、酒會、升旗、計算、勾選狀態、自訂條件），僅保留 `rowNumber` 在 `matchedSectionRowNumbers` 內的 `section` 列（或是本身直接命中關鍵字搜尋且無其他限定條件的 `section` 列）。
     - 若完全無篩選條件，顯示全部 `section` 列與 `record` 列。
  4. 當**啟用欄位排序 (`sort?.column != null`)** 時：
     - 為了讓跨組別欄位排序（例如按姓名、年齡、職位排序）結果清晰一致，檢視列表僅排列命中的 `record` 列（隱藏分組標題列），並在工具列顯示提示「檢視排序中（匯出仍維持原檔組別與列順序）」。
     - 排序採用數值感知與繁體中文語系比較：`Intl.Collator('zh-Hant', { numeric: true, sensitivity: 'base' })`，空值永遠排在最後。

### 5.3 選取與批次操作 (Selection Workflow)

- 以 `selectedRowNumbers: Set<number>` 追蹤已勾選的 `record` 列（以原檔 `rowNumber` 為唯一鍵，不受篩選或排序切換影響）。
- 支援的操作：
  1. **單列勾選**：點擊每筆 `record` 列最左側 Checkbox。
  2. **分組標題列整組勾選**：每個 `section` 標題列左側亦有 Checkbox（顯示全選 / 部分選取 indeterminate / 未選狀態），點擊後可一鍵將該界別或組別底下（在目前篩選範圍內）的所有成員加入或移出選取。
  3. **左側界別/組別面板勾選**：在左側導覽樹的界別或組別旁直接點擊 Checkbox 整批勾選/取消。
  4. **批次工具列按鈕**：
     - `全選目前篩選 (N)`：將目前所有符合篩選條件的 `record` 列加入已選集合。
     - `取消目前篩選`：從已選集合移除目前篩選結果。
     - `反選目前篩選`：對目前篩選結果進行勾選狀態反轉。
     - `清空全部已選 (M)`：一鍵清除所有勾選。

### 5.4 萬筆虛擬捲動資料表 (`src/components/VirtualizedDbTable.tsx`)

- 為確保 10,000+ 筆資料在切換篩選、排序、全選或捲動時維持 60fps：
  - 使用固定高度捲動容器（如 `height: 620px`, `overflow: auto`）搭配 Sticky 表頭 (`thead`)。
  - 根據 `scrollTop`、固定列高（`ROW_HEIGHT = 40px`）與可視高度計算 `[startIndex, endIndex]`（外加前後 `overscan = 15` 列緩衝），透過上下 Spacer `<tr>` 撐開總捲動高度。
  - 支援點擊任意欄位標頭切換排序：`原檔順序` -> `升冪 (▲)` -> `降冪 (▼)` -> `原檔順序`。
  - 支援「快速跳轉至組別/列號」與鍵盤捲動。

### 5.5 原結構 Excel 匯出引擎 (`src/lib/excelExporter.ts`) 與 IPC

1. **匯出函式 `buildExportWorkbookBuffer(workbook: ParsedCrmWorkbook, options: ExportCrmOptions)`**：
   - `options.mode`: `'filtered'`（匯出目前篩選命中的記錄）或 `'selected'`（僅匯出已勾選的記錄）。
   - **決定匯出的 `DB` 列**：
     - 取得目標 `record` 列集合（依 `rowNumber` 升冪排列，**永遠維持原檔列順序**）。
     - 計算這批目標 `record` 列的 `ancestorSectionRows` 聯集，得到所有必須保留的上層分組標題列（含界別 `MA000`、母組 `MA001`、子組 `MA001-01`）。
     - 將這些保留的 `section` 列與目標 `record` 列合併，並嚴格依原檔 `rowNumber` 由小到大排序。
   - **建立 `DB` 工作表**：
     - 第 1 列為完整 50 個 `workbook.dbHeaders`。
     - 第 2 列起依序寫入上述排序後的 `section` 與 `record` 列（全部 50 欄完整輸出）。
     - 針對每一列輸出的 `section` 列（位於輸出表的 0-based 列索引 `r`），在 `ws['!merges']` 加入 `{ s: { c: 6, r }, e: { c: 7, r } }`（合併第 7、8 欄 `姓名` 與 `職位`），與原檔 `test.xlsx` 格式完全一致。
   - **建立 `組別` 與 `界別` 工作表**：
     - 使用原檔保存的 `rawGroupRows` 與 `rawSectorRows`（若無則由 `groups` / `sectors` 重建）完整寫入 `組別` 與 `界別` 工作表。
   - 使用 `XLSX.write(wb, { type: 'array', bookType: 'xlsx' })` 產出 `Uint8Array`。

2. **匯出結果摘要與檔案儲存流程**：
   - 頂部導覽列與 DB 工具列提供醒目的「匯出 Excel」按鈕與匯出模式選擇（或彈出確認列）：
     - 顯示預計匯出的：`名單筆數 (N 筆)`、`保留分組標題列數 (S 列)`、`完整 50 欄位 + 組別/界別工作表`。
     - 預設檔名建議：`{原檔名去副檔名}_篩選匯出_{YYYYMMDD_HHmm}.xlsx`。
   - **Electron 模式**：新增 IPC `excel:save-file`，呼叫 `dialog.showSaveDialog` 讓使用者選擇儲存路徑並以 `fs.writeFile` 寫入磁碟，成功後在介面顯示儲存路徑提示。
   - **瀏覽器 Fallback 模式**：產生 `Blob` 並觸發 `<a download>` 下載。

---

## 6. 留給實作者的細節 (Left to Builder)

- 虛擬捲動元件的具體緩衝列數（建議 12～20 列）與列高微調。
- 左側界別/組別導覽面板的收合動畫與寬度配置（建議預設寬度 `280px`，可一鍵收合以最大化表格可視寬度）。
- 自訂條件篩選器的新增/刪除列 UI 細節與空白狀態提示文案。
