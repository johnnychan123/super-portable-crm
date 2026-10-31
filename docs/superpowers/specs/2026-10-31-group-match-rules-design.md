# 外部篩選 Excel 預設比對依據與各組別專屬比對規則 — 設計文件

## 1. 目標與背景 (Intent & Why)
在實際 CRM 名單中，不同組別（如 `MA001` vs `MZ001`）的資料完整度可能不同（例如部分組別有完整 `PID`，而部分組別未填 `PID` 需改用 `組別 + 姓名` 或 `姓名` 比對）。
本功能讓使用者能**先設定一個全域預設的比對依據**，並可**針對個別組別設定不同的專屬比對依據**，同時保留單一匯入檔案覆寫比對依據的彈性。

## 2. 目標與非目標 (Goals, Non-Goals, Anti-Goals)
- **Goals:**
  - 支援設定「全域預設比對依據」（如 `auto`、`Pf.ID`、`PID+組別`、`PID`、`組別+姓名`、`姓名`、`中文名(沒有稱謂)` 或共同欄位）。
  - 支援新增/編輯/刪除多筆「指定組別 → 專屬比對依據」例外規則列，並提供組別搜尋過濾以便從 558 個組別或外部檔組別中快速選取。
  - 每份匯入的外部篩選 Excel 預設採用 `'global'`（跟隨全域預設 + 組別專屬規則），亦可在檔案卡片下拉選單單獨覆寫為特定比對依據。
  - 比對引擎 (`matchExternalFilterToDb`) 對設定了專屬規則的組別，嚴格使用該組別指定的欄位建立鍵值；未設定專屬規則的組別則使用預設比對依據。
- **Non-Goals:**
  - 不修改主 DB 原始資料內容。
- **Anti-Goals:**
  - 當該組別指定的比對欄位在某列為空白時，嚴格不自動退回其他欄位、亦絕不讓空白與空白碰撞。

## 3. 方案比較 (Approaches)
- **方案 A（已選用）：在「外部篩選 Excel」面板內直接整合「全域比對規則設定區（預設比對依據 + 多選組別/Prefix 快速選取規則卡片）」**
  - 優點：在同一面板即可輸入 Prefix（如 `MZ`、`MA0`）一鍵批次勾選/取消所有符合前綴的組別，同時檢視已選組別標籤（Chips）與核取方塊清單進行個別微調，並即時反映各檔案命中筆數。
- **方案 B：透過獨立 Modal 或浮動選單管理組別規則**
  - 缺點：多一層視窗切換，調整 Prefix 與多選組別時無法即時對照下方各檔案的命中筆數變化。

## 4. 核心設計細節 (Technical Details)
1. **資料結構 (`src/lib/externalFilterMatcher.ts`)**：
   - `GroupMatchOverride`:
     - `id: string`
     - `groupCodes: string[]`（支援一條規則同時綁定多個組別，例如 `['MZ001', 'MZ002', ...]`；同時相容舊欄位 `groupCode?: string`）
     - `prefixes?: string[]`（可選：若指定前綴亦可直接匹配該前綴開頭的組別）
     - `matchKey: string`
   - `filterGroupsByPrefix(options: GroupOptionForOverride[], query: string): GroupOptionForOverride[]`：
     - 當使用者輸入 Prefix（如 `MZ` 或 `MA01`）時，優先比對 `code.toUpperCase().startsWith(prefix.toUpperCase())`（若輸入中文則同時支援名稱子字串搜尋），方便一鍵批次勾選或取消該前綴的所有組別。
   - `matchExternalFilterToDb(sheet, dbRows, matchKeyOrMode, config)`:
     - 當檔案模式為 `'global'` 時，將每條 `GroupMatchOverride` 內的 `groupCodes`（及 `groupCode`）全部展開映射到該規則的 `matchKey`（若同組別在多條規則出現，後者覆寫前者），另外若設定了 `prefixes` 則依前綴匹配。
2. **UI 互動 (`src/App.tsx`)**：
   - 每條「組別專屬對比規則」卡片支援：
     - **設定該規則的對比依據 (`matchKey`)**。
     - **Prefix 快速選取列**：輸入框（如輸入 `MZ`、`MA`），即時計算符合該前綴的組別數，提供 **「+ 全選此前綴組別 (N)」** 與 **「- 取消此前綴組別」** 按鈕（在輸入框按 Enter 亦可直接將符合前綴的組別全數加入）。
     - **已選組別標籤 (Chips) 與微調**：顯示目前已選取的組別代碼與名稱（可點擊 `×` 移除單一組別，或點擊「清空已選」），下方提供緊湊的多選 Checkbox 網格供單獨勾選/取消勾選。

## 5. 自訂複合欄對比依據 (Custom Composite Match Key Columns)
1. **編碼與解析 (`src/lib/externalFilterMatcher.ts`)**：
   - 新增 `CUSTOM_MATCH_KEY_PREFIX = 'custom:'` 與分隔符 `'\u001f'`：
     - `isCustomMatchKey(matchKey: string): boolean`：判斷是否為 `'custom'` 或以 `'custom:'` 開頭。
     - `encodeCustomMatchKey(columns: ReadonlyArray<string>): string`：將使用者勾選的多個欄位（去除空白與重複）編碼為 `custom:<col1>\u001f<col2>...`（若尚未勾選則為 `'custom:'`）。
     - `parseCustomMatchKeyColumns(matchKey: string): string[]`：從 `custom:...` 解析出欄位陣列。
   - `resolveMatchKeyColumns(matchKey, sheet, dbRows)`：
     - 若 `isCustomMatchKey(matchKey)` 為真，解析出指定欄位陣列 `cols`；僅當 `cols.length > 0` 且所有欄位皆存在於 `sheet.sharedHeaders` 時回傳 `cols`，否則回傳 `[]`（確保外部檔缺漏任一自訂複合欄時嚴格不誤配）。
   - 新增 `prioritizeMatchColumns(headers: ReadonlyArray<string>): string[]`：
     - 將常用識別欄位（`組別`、`PID`、`Pf.ID`、`姓名`、`中文名(沒有稱謂)`、`英文名(沒有稱謂)`、`手提電話`、`電郵`、`公司名稱`）優先排在最前方，其餘欄位依原順序排列，方便使用者一鍵快速勾選。
2. **UI 互動（精簡主畫面 + 彈出視窗 Modal）(`src/App.tsx`)**：
   - 主畫面的「外部篩選 Excel」面板保持精簡單行列表，不在主畫面直接展開龐大的組別清單或自訂複合欄勾選面板：
     - **預設對比依據列**：當選為自訂複合欄時，旁側顯示目前已選複合欄摘要標籤（如 `組別 + 姓名 + 手提電話`）與 **「設定複合欄...」** 按鈕。
     - **各組別專屬規則列（單行精簡卡片）**：顯示 `規則 #N`、已套用組別摘要（如 `已選 12 個組別：MZ001, MZ002, MZ003...`）與 **「選取組別...」** 按鈕、對比依據下拉選單（若為自訂複合欄則附帶摘要標籤與 **「設定複合欄...」** 按鈕）、以及刪除規則按鈕。新增規則時自動開啟「選取組別 Modal」。
     - **單一篩選檔卡片**：當單檔對比依據選為自訂複合欄時，旁側顯示複合欄摘要標籤與 **「設定複合欄...」** 按鈕。
   - **Modal 1：套用組別選取視窗 (`GroupPickerModalState`)**：
     - 採用暫存草稿狀態（`draftGroupCodes`、`prefixInput`），使用者在視窗內可輸入 Prefix（按 Enter 或點按鈕批次全選／取消前綴組別）、點擊已選組別 Chips 的 `×` 移除、或在寬敞的多選核取網格微調；點擊 **「確認套用」** 才正式寫回規則並重算篩選結果，點擊 **「取消」** 或背景遮罩則不影響原設定。
   - **Modal 2：自訂複合欄設定視窗 (`CustomCompositeModalState`)**：
     - 支援三種來源目標（`'default'` 全域預設、`{ type: 'rule', ruleId }` 組別規則、`{ type: 'file', fileId }` 單一篩選檔）。
     - 採用暫存草稿狀態（`draftColumns`、`searchQuery`），支援搜尋過濾欄位、依優先順序勾選／取消多個欄位、檢視與移除已選欄位組合順序；點擊 **「確認套用」**（需至少勾選 1 個欄位）才正式更新對應的 `matchKey` 並重算，點擊 **「取消」** 則保留開啟視窗前的設定。


