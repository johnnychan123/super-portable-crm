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
- **方案 A（已選用）：在「外部篩選 Excel」面板內直接整合「全域比對規則設定區（預設比對依據 + 組別例外規則列）」**
  - 優點：在同一面板即可一覽全域預設依據、各組別例外規則與各檔案命中數變化，修改規則時即時重新計算命中結果。
- **方案 B：透過獨立 Modal 管理組別規則**
  - 缺點：多一層視窗切換，調整規則時無法即時對照下方各檔案的命中筆數變化。

## 4. 核心設計細節 (Technical Details)
1. **資料結構 (`src/lib/externalFilterMatcher.ts`)**：
   - `GroupMatchOverride`: `{ id: string; groupCode: string; matchKey: string }`
   - `resolveMatchKeyToColumns(matchKey: string, sheet: ParsedExternalFilterSheet, dbRows: DbRowItem[]): string[]`
   - `matchExternalFilterToDb(sheet, dbRows, matchKeyOrMode, options?: { defaultMatchKey?: string; groupOverrides?: Record<string, string> })`:
     - 當檔案模式為 `'global'`（或傳入 `groupOverrides`）時：
       - 若某列（DB 列或外部篩選檔列）的 `組別` 存在於 `groupOverrides`（忽略大小寫空白差異），則該列僅使用 `groupOverrides[groupCode]` 對應的欄位組合建立比對鍵值。
       - 若該列的 `組別` 未設定專屬規則（或該列無 `組別`），則使用預設比對依據（`defaultMatchKey`）。
       - 為避免不同規則的鍵值互相干擾，每種規則生成的複合鍵值加上規則欄位前綴（`[cols.join('+')]:value`），且只在相同規則下比對。
2. **UI 互動 (`src/App.tsx`)**：
   - 在 `showExternalFilterPanel` 內新增「比對規則設定（預設比對依據 + 各組別專屬比對依據）」區塊：
     - **預設比對依據下拉選單**：預設為 `auto`（自動偵測），可改為 `Pf.ID`、`PID+組別`、`PID`、`組別+姓名`、`姓名`、`中文名(沒有稱謂)` 或任意共同欄位。
     - **「+ 新增組別比對規則」按鈕**：點擊新增一列規則，包含：
       - 組別搜尋輸入框 + 組別選擇下拉選單（優先列出目前已匯入篩選檔內出現過的組別，以及主 DB 全部組別，顯示 `MA001 行政長官辦公室 (篩選檔內 N 筆)`）。
       - 該組別專屬比對依據下拉選單。
       - 刪除規則按鈕。
   - 當 `defaultMatchKey` 或 `groupMatchOverrides` 變更時，所有處於 `'global'`（跟隨全域規則）模式的已匯入篩選檔會自動即時重新計算 `matchResult`。
