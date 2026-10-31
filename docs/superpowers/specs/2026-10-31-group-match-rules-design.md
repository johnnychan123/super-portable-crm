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
