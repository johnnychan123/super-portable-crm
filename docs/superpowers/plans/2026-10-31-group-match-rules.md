# 外部篩選 Excel 預設比對依據與各組別專屬比對規則 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在外部篩選 Excel 比對功能中，支援先設定一個全域預設比對依據，並可針對各組別（`組別` 代碼）設定專屬比對依據，且保留單一檔案覆寫能力。

**Architecture:** 擴充 `src/lib/externalFilterMatcher.ts` 的 `matchExternalFilterToDb` 以支援 `ExternalFilterMatchConfig`（包含 `defaultMatchKey` 與 `groupOverrides`），依每筆記錄的 `組別` 動態決定比對欄位並嚴格防止空白碰撞；在 `src/App.tsx` 的外部篩選面板中加入「全域預設比對依據 + 各組別專屬比對規則列（含組別搜尋）」設定區，變更時即時重算所有跟隨全域規則的篩選檔命中結果。

**Tech Stack:** TypeScript, React 19, SheetJS (`xlsx`), Bun test

**Spec:** `docs/superpowers/specs/2026-10-31-group-match-rules-design.md`

## Global Constraints

- 當某組別設定了專屬比對依據時，該組別的記錄必須嚴格使用該組別指定的欄位比對；若該欄位在某列為空白，絕不自動退回其他欄位，亦絕不與其他空白儲存格碰撞。
- 未設定專屬規則的組別（或無組別欄位的記錄），使用全域預設比對依據（`defaultMatchKey`）。
- 每份匯入的篩選 Excel 預設使用 `'global'`（跟隨全域預設 + 組別規則），亦可在檔案卡片單獨覆寫為特定比對依據。

## Review Focus

- **Per-group strict isolation:** 當組別 `MA001` 使用預設 `PID+組別`、而組別 `MZ001` 設定專屬規則 `組別+姓名` 時，`MZ001` 即使沒有 `PID` 也能透過 `組別+姓名` 命中，而 `MA001` 若某列缺少 `PID` 則不會誤用 `組別+姓名` 命中。
- **Cross-rule key namespace isolation:** 不同規則產生的複合鍵值必須帶有欄位組合命名空間前綴（例如 `PID+組別::...` vs `組別+姓名::...`），避免單一欄值意外跨規則碰撞。
- **Unavailable column in external file for a group override:** 若某組別設定了專屬比對欄位（如 `Pf.ID`），但某份匯入的篩選 Excel 刪除了 `Pf.ID` 欄，該組別在該檔案應安全返回 0 命中而不拋出例外。
- **Live recomputation on rule edit:** 修改全域預設比對依據或新增/修改/刪除組別專屬規則時，所有 `selectedMatchKey === 'global'` 的已匯入檔案必須立即更新 `matchResult`，而單獨覆寫（`selectedMatchKey !== 'global'`）的檔案不受影響。

---

### Task 1: Group-Aware External Filter Matching Engine

**Files:**
- Modify: `src/lib/externalFilterMatcher.ts`
- Test: `src/__tests__/externalFilterMatcher.test.ts`

**Interfaces:**
- Produces:
  - `GroupMatchOverride`: `{ id: string; groupCode: string; matchKey: string }`
  - `ExternalFilterMatchConfig`: `{ defaultMatchKey: string; groupOverrides?: ReadonlyArray<GroupMatchOverride> }`
  - `getGlobalMatchKeyOptions(dbHeaders: string[]): ExternalMatchKeyOption[]`
  - `resolveMatchKeyColumns(matchKey: string, sheet: ParsedExternalFilterSheet, dbRows: DbRowItem[]): string[]`
  - `matchExternalFilterToDb(sheet: ParsedExternalFilterSheet, dbRows: DbRowItem[], matchKey?: string, config?: ExternalFilterMatchConfig): ExternalFilterMatchResult`

- [ ] **Step 1: Write the failing test in `src/__tests__/externalFilterMatcher.test.ts`**

```ts
  test('supports defaultMatchKey plus per-group matchKey overrides with strict blank non-fallback', () => {
    const buf = makeWorkbookBuffer('DB', [
      ['組別', 'PID', '姓名', '桌號'],
      ['MA001', '101', '陳大文', 'A1'], // MA001 uses default (PID+組別) -> matches row 3
      ['MA001', '', '李小明', 'A2'],    // MA001 missing PID -> strict no-fallback -> should NOT match row 4
      ['MZ001', '', '張三', 'Z1'],      // MZ001 overridden to use 組別+姓名 -> matches row 5 even without PID
    ]);

    const extendedDbRows: DbRowItem[] = [
      ...dbRows.slice(0, 3),
      {
        rowNumber: 5,
        kind: 'record',
        values: { 'Pf.ID': '', PID: '', 組別: 'MZ001', 姓名: '張三', '中文名(沒有稱謂)': '張三' },
        ancestorSectionRows: [2],
      },
    ];

    const parsed = parseExternalFilterWorkbook(buf, 'mixed_groups.xlsx', dbHeaders);
    const res = matchExternalFilterToDb(parsed, extendedDbRows, 'global', {
      defaultMatchKey: 'PID+組別',
      groupOverrides: [{ id: 'g1', groupCode: 'MZ001', matchKey: '組別+姓名' }],
    });

    expect(Array.from(res.matchedDbRowNumbers)).toEqual([3, 5]);
    expect(res.matchedExternalRowCount).toBe(2);
    expect(res.unmatchedExternalRowCount).toBe(1);
    expect(res.extraValuesByDbRow.get(3)).toEqual({ 桌號: 'A1' });
    expect(res.extraValuesByDbRow.get(5)).toEqual({ 桌號: 'Z1' });
  });
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun test src/__tests__/externalFilterMatcher.test.ts`
Expected: FAIL (`Expected: [3, 5], Received: []`)

- [ ] **Step 3: Implement `GroupMatchOverride`, `ExternalFilterMatchConfig`, `getGlobalMatchKeyOptions`, and group-aware `matchExternalFilterToDb` in `src/lib/externalFilterMatcher.ts`**

Update `src/lib/externalFilterMatcher.ts`:
- Export `GroupMatchOverride` and `ExternalFilterMatchConfig`.
- Export `getGlobalMatchKeyOptions(dbHeaders: string[]): ExternalMatchKeyOption[]` returning `'auto'`, `'Pf.ID'`, `'PID+組別'`, `'PID'`, `'組別+姓名'`, `'姓名'`, `'中文名(沒有稱謂)'`, plus remaining `dbHeaders`.
- Export `resolveMatchKeyColumns(matchKey: string, sheet: ParsedExternalFilterSheet, dbRows: DbRowItem[]): string[]` which resolves `'auto'`, `'PID+組別'`, `'組別+姓名'`, or a single column name, and returns `[]` if any required column is not in `sheet.sharedHeaders`.
- Update `getAvailableMatchKeys(sheet, dbRows)` to include `{ key: 'global', label: '跟隨全域規則（預設 + 組別設定）', columns: autoCols }` as the first option.
- Update `matchExternalFilterToDb(sheet, dbRows, matchKey = 'auto', config?: ExternalFilterMatchConfig)`:
  - When `matchKey === 'global'` (or `config` is used with `'global'`): resolve `defaultCols = resolveMatchKeyColumns(config?.defaultMatchKey ?? 'auto', sheet, dbRows)` and map each `override` in `config?.groupOverrides ?? []` with non-empty `groupCode` to `resolveMatchKeyColumns(override.matchKey, sheet, dbRows)`.
  - For each row, select `cols = groupOverrideCols.get(normalizedGroupCode) ?? defaultCols`, build key `${cols.join('+')}::${compositeKey}`, and match external rows against DB record rows.

- [ ] **Step 4: Run test to verify it passes**

Run: `bun test src/__tests__/externalFilterMatcher.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/lib/externalFilterMatcher.ts src/__tests__/externalFilterMatcher.test.ts
git commit -m "feat(matcher): support default match key and per-group match key overrides"
```

---

### Task 2: Global Default & Per-Group Match Rules UI in `src/App.tsx`

**Files:**
- Modify: `src/App.tsx`
- Test: `src/__tests__/app.test.ts`

**Interfaces:**
- Consumes: `GroupMatchOverride`, `ExternalFilterMatchConfig`, `getGlobalMatchKeyOptions`, `matchExternalFilterToDb` from `src/lib/externalFilterMatcher.ts`
- Produces:
  - `recomputeFilesWithGlobalRules(files: ImportedFilterFile[], dbRows: DbRowItem[], config: ExternalFilterMatchConfig): ImportedFilterFile[]` exported from `src/App.tsx`
  - UI controls in `showExternalFilterPanel` for selecting `defaultExternalMatchKey` and adding/updating/removing `groupMatchOverrides` with group search filtering.

- [ ] **Step 1: Write the failing test in `src/__tests__/app.test.ts`**

```ts
  test('recomputeFilesWithGlobalRules updates files using global mode while preserving single-file overrides', () => {
    const dbHeaders = ['組別', 'PID', '姓名'];
    const dbRows: DbRowItem[] = [
      {
        rowNumber: 3,
        kind: 'record',
        values: { 組別: 'MZ001', PID: '', 姓名: '張三' },
        ancestorSectionRows: [],
      },
    ];
    const sheet: ParsedExternalFilterSheet = {
      fileName: 'ext.xlsx',
      sheetName: 'DB',
      headers: ['組別', 'PID', '姓名'],
      rows: [{ 組別: 'MZ001', PID: '', 姓名: '張三' }],
      sharedHeaders: ['組別', 'PID', '姓名'],
      extraHeaders: [],
    };
    const initialFiles: ImportedFilterFile[] = [
      {
        id: 'f-global',
        fileName: 'global.xlsx',
        sheet,
        availableMatchKeys: [],
        selectedMatchKey: 'global',
        matchResult: matchExternalFilterToDb(sheet, dbRows, 'global', {
          defaultMatchKey: 'PID+組別',
          groupOverrides: [],
        }),
      },
      {
        id: 'f-override',
        fileName: 'override.xlsx',
        sheet,
        availableMatchKeys: [],
        selectedMatchKey: 'PID+組別',
        matchResult: matchExternalFilterToDb(sheet, dbRows, 'PID+組別'),
      },
    ];

    expect(initialFiles[0]?.matchResult.matchedDbRowNumbers.size).toBe(0);

    const updated = recomputeFilesWithGlobalRules(initialFiles, dbRows, {
      defaultMatchKey: 'PID+組別',
      groupOverrides: [{ id: 'g1', groupCode: 'MZ001', matchKey: '組別+姓名' }],
    });

    expect(Array.from(updated[0]?.matchResult.matchedDbRowNumbers ?? [])).toEqual([3]);
    expect(updated[1]?.matchResult.matchedDbRowNumbers.size).toBe(0);
  });
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun test src/__tests__/app.test.ts`
Expected: FAIL (`Export named 'recomputeFilesWithGlobalRules' not found`)

- [ ] **Step 3: Implement `recomputeFilesWithGlobalRules` and the Global + Per-Group Match Rules UI in `src/App.tsx`**

In `src/App.tsx`:
- Export `recomputeFilesWithGlobalRules(files, dbRows, config)` which recalculates `matchResult` for every file whose `selectedMatchKey === 'global'`.
- Add state:
  - `defaultExternalMatchKey: string` (`'auto'`)
  - `groupMatchOverrides: GroupMatchOverride[]` (`[]`)
  - `groupOverrideSearch: string` (`''`)
- Set newly imported files in `handleAddExternalFilterBuffers` to `selectedMatchKey: 'global'` and compute `matchResult` with `{ defaultMatchKey: defaultExternalMatchKey, groupOverrides: groupMatchOverrides }`.
- Add handlers for changing `defaultExternalMatchKey`, adding a `GroupMatchOverride`, updating a `GroupMatchOverride` (`groupCode` or `matchKey`), and removing a `GroupMatchOverride`, each calling `recomputeFilesWithGlobalRules`.
- In `showExternalFilterPanel`, render the "比對依據設定（預設比對依據 + 各組別專屬比對依據）" card with:
  - Default match key selector (`getGlobalMatchKeyOptions(workbook.dbHeaders)`)
  - `+ 新增組別專屬比對依據` button
  - Search input for filtering groups + list of `groupMatchOverrides` rows (`組別` selector showing external file counts + group name, `比對依據` selector, and delete button).

- [ ] **Step 4: Run test to verify it passes**

Run: `bun run typecheck && bun test && bun run build`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/App.tsx src/__tests__/app.test.ts docs/superpowers/specs/2026-10-31-group-match-rules-design.md docs/superpowers/plans/2026-10-31-group-match-rules.md
git commit -m "feat(ui): add default match key and per-group match key override rules UI"
```
