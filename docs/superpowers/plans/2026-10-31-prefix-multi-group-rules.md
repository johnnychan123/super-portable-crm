# 組別專屬對比規則多選與 Prefix 快速選取 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 讓每條「組別專屬對比規則」支援同時多選多個組別（`groupCodes: string[]`），並提供輸入 Prefix（如 `MZ`、`MA01`）一鍵批次勾選/取消符合前綴的組別及顯示已選組別標籤供微調。

**Architecture:** 擴充 `src/lib/externalFilterMatcher.ts` 的 `GroupMatchOverride` 支援 `groupCodes?: string[]`（向下相容 `groupCode?: string`），新增 `filterGroupsByPrefix` 與 `applyPrefixGroupSelection` 輔助函式；在 `src/App.tsx` 的組別規則卡片中提供 Prefix 輸入框、一鍵勾選/取消前綴組別按鈕、已選組別標籤（Chips）與多選 Checkbox 網格。

**Tech Stack:** TypeScript, React 19, Bun test

**Spec:** `docs/superpowers/specs/2026-10-31-group-match-rules-design.md`

## Global Constraints

- 一條 `GroupMatchOverride` 規則可包含多個組別代碼（`groupCodes: string[]`），所有包含在內的組別皆嚴格使用該規則指定的 `matchKey` 進行比對（空白不退回、不碰撞）。
- 向下相容單一 `groupCode?: string` 屬性。
- 輸入 Prefix 時（忽略大小寫與前後空白），以組別代碼前綴（`code.toUpperCase().startsWith(prefix)`）進行快速篩選，若輸入非純英數代碼亦支援比對組別中文名稱；支援一鍵將符合前綴的所有組別加入或移出該規則的 `groupCodes`。

## Review Focus

- **Multi-group rule matching:** 當同一條規則設定 `groupCodes: ['MZ001', 'MZ002']` → `組別+姓名` 時，兩個組別皆使用 `組別+姓名` 比對，而未列入的 `MA001` 仍使用 `defaultMatchKey`。
- **Prefix matching accuracy:** 輸入前綴 `MZ` 時，`filterGroupsByPrefix` 只匹配 `MZ` 開頭的組別（如 `MZ001`、`MZ002`），不會誤把 `MA001` 納入；按「勾選此前綴」時自動去重保留既有已選的其他組別（如原本已選 `MA009`，再選前綴 `MZ` 後結果為 `['MA009', 'MZ001', 'MZ002']`）。
- **Deselect by prefix:** 按「取消此前綴」時，僅從已選清單移除符合該前綴的組別，保留其他已選組別。
- **Empty `groupCodes` in a rule:** 當某條規則尚未勾選任何組別（`groupCodes: []`）時，不會影響任何組別的預設比對行為。

---

### Task 1: Multi-Group Override Matching & Prefix Selection Helpers

**Files:**
- Modify: `src/lib/externalFilterMatcher.ts`
- Test: `src/__tests__/externalFilterMatcher.test.ts`

**Interfaces:**
- Produces:
  - `GroupMatchOverride`: `{ id: string; groupCodes?: string[]; groupCode?: string; matchKey: string }`
  - `filterGroupsByPrefix(options: GroupOptionForOverride[], prefixOrQuery: string): GroupOptionForOverride[]`
  - `applyPrefixGroupSelection(currentCodes: string[], options: GroupOptionForOverride[], prefixOrQuery: string, mode: 'select' | 'deselect'): string[]`

- [ ] **Step 1: Write the failing test in `src/__tests__/externalFilterMatcher.test.ts`**

```ts
  test('supports multi-group groupCodes in a single rule and prefix batch selection/deselection', () => {
    const buf = makeWorkbookBuffer('DB', [
      ['組別', 'PID', '姓名', '桌號'],
      ['MA001', '', '陳大文', 'A1'], // MA001 uses default PID+組別 -> missing PID -> not matched
      ['MZ001', '', '李小明', 'Z1'], // MZ001 in multi-group rule (組別+姓名) -> matched
      ['MZ002', '', '張三', 'Z2'],   // MZ002 in multi-group rule (組別+姓名) -> matched
    ]);

    const multiGroupDbRows: DbRowItem[] = [
      {
        rowNumber: 3,
        kind: 'record',
        values: { 組別: 'MA001', PID: 'P001', 姓名: '陳大文' },
        ancestorSectionRows: [],
      },
      {
        rowNumber: 4,
        kind: 'record',
        values: { 組別: 'MZ001', PID: '', 姓名: '李小明' },
        ancestorSectionRows: [],
      },
      {
        rowNumber: 5,
        kind: 'record',
        values: { 組別: 'MZ002', PID: '', 姓名: '張三' },
        ancestorSectionRows: [],
      },
    ];

    const groupOpts = [
      { code: 'MA001', nameZh: '行政長官辦公室', category: 'A', externalCount: 1 },
      { code: 'MZ001', nameZh: '基本法委員會', category: 'B', externalCount: 1 },
      { code: 'MZ002', nameZh: '選舉委員會', category: 'B', externalCount: 1 },
    ];

    expect(filterGroupsByPrefix(groupOpts, 'mz').map((g) => g.code)).toEqual(['MZ001', 'MZ002']);

    const selectedAfterPrefix = applyPrefixGroupSelection(['MA009'], groupOpts, 'MZ', 'select');
    expect(selectedAfterPrefix).toEqual(['MA009', 'MZ001', 'MZ002']);

    const selectedAfterDeselect = applyPrefixGroupSelection(
      selectedAfterPrefix,
      groupOpts,
      'MZ001',
      'deselect',
    );
    expect(selectedAfterDeselect).toEqual(['MA009', 'MZ002']);

    const parsed = parseExternalFilterWorkbook(buf, 'multi_mz.xlsx', dbHeaders);
    const res = matchExternalFilterToDb(parsed, multiGroupDbRows, 'global', {
      defaultMatchKey: 'PID+組別',
      groupOverrides: [
        { id: 'rule-mz', groupCodes: ['MZ001', 'MZ002'], matchKey: '組別+姓名' },
      ],
    });

    expect(Array.from(res.matchedDbRowNumbers)).toEqual([4, 5]);
  });
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun test src/__tests__/externalFilterMatcher.test.ts`
Expected: FAIL (`Export named 'filterGroupsByPrefix' not found`)

- [ ] **Step 3: Implement `groupCodes` support, `filterGroupsByPrefix`, and `applyPrefixGroupSelection` in `src/lib/externalFilterMatcher.ts`**

- Update `GroupMatchOverride`:
  ```ts
  export interface GroupMatchOverride {
    id: string;
    groupCodes?: string[];
    groupCode?: string;
    matchKey: string;
  }
  ```
- In `matchExternalFilterToDb`: iterate over `[...(override.groupCodes ?? []), ...(override.groupCode ? [override.groupCode] : [])]` and register each normalized non-empty code into `groupOverrideColumnsMap`.
- Implement `filterGroupsByPrefix(options, prefixOrQuery)`:
  - Trim and uppercase query; if empty, return `options`.
  - Match items where `item.code.toUpperCase().startsWith(normQuery)` or `item.nameZh.toLowerCase().includes(rawQuery.toLowerCase())`.
- Implement `applyPrefixGroupSelection(currentCodes, options, prefixOrQuery, mode)`:
  - Find matched codes via `filterGroupsByPrefix(options, prefixOrQuery)`.
  - In `'select'` mode, append any matched code not already in `currentCodes` (case-insensitive check).
  - In `'deselect'` mode, remove any code in `currentCodes` that matches the target set.

- [ ] **Step 4: Run test to verify it passes**

Run: `bun test src/__tests__/externalFilterMatcher.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/lib/externalFilterMatcher.ts src/__tests__/externalFilterMatcher.test.ts
git commit -m "feat(matcher): support multi-group overrides and prefix group selection helpers"
```

---

### Task 2: Prefix Batch Selection & Multi-Select Group Override UI in `src/App.tsx`

**Files:**
- Modify: `src/App.tsx`
- Test: `src/__tests__/app.test.ts`

**Interfaces:**
- Consumes: `filterGroupsByPrefix`, `applyPrefixGroupSelection`, `GroupMatchOverride` from `src/lib/externalFilterMatcher.ts`
- Produces:
  - Per-rule prefix input (`rulePrefixInputs: Record<string, string>`), prefix batch select/deselect buttons, selected group chips (`Badge` with `×` remove), and multi-select group checkbox grid in `src/App.tsx`.

- [ ] **Step 1: Write the failing test in `src/__tests__/app.test.ts`**

```ts
  test('recomputeFilesWithGlobalRules supports multi-group groupCodes overrides', () => {
    const dbRows: DbRowItem[] = [
      {
        rowNumber: 4,
        kind: 'record',
        values: { 組別: 'MZ001', PID: '', 姓名: '李小明' },
        ancestorSectionRows: [],
      },
      {
        rowNumber: 5,
        kind: 'record',
        values: { 組別: 'MZ002', PID: '', 姓名: '張三' },
        ancestorSectionRows: [],
      },
    ];
    const sheet: ParsedExternalFilterSheet = {
      fileName: 'ext.xlsx',
      sheetName: 'DB',
      headers: ['組別', 'PID', '姓名'],
      rows: [
        { 組別: 'MZ001', PID: '', 姓名: '李小明' },
        { 組別: 'MZ002', PID: '', 姓名: '張三' },
      ],
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
    ];

    const updated = recomputeFilesWithGlobalRules(initialFiles, dbRows, {
      defaultMatchKey: 'PID+組別',
      groupOverrides: [{ id: 'g1', groupCodes: ['MZ001', 'MZ002'], matchKey: '組別+姓名' }],
    });

    expect(Array.from(updated[0]?.matchResult.matchedDbRowNumbers ?? [])).toEqual([4, 5]);
  });
```

- [ ] **Step 2: Run test to verify it fails or update UI**

Run: `bun test src/__tests__/app.test.ts`

- [ ] **Step 3: Implement multi-group & prefix selection UI in `src/App.tsx`**

In `src/App.tsx`:
- Update `handleAddGroupMatchOverride` to initialize `groupCodes: candidateGroup ? [candidateGroup] : []`.
- Add `rulePrefixInputs: Record<string, string>` state so each rule card has its own Prefix input box (e.g. `MZ`, `MA01`).
- Add handlers:
  - `handleToggleGroupInOverride(ruleId: string, groupCode: string)`
  - `handlePrefixBatchSelectInOverride(ruleId: string, prefix: string, mode: 'select' | 'deselect')`
  - `handleClearGroupsInOverride(ruleId: string)`
- In each rule card inside `showExternalFilterPanel`:
  - Top row: Match key selector (`使用對比依據`), badge showing `已選 N 個組別`, `清空已選組別` button, and delete rule button.
  - Second row (Prefix bar): Input `輸入組別 Prefix 快速選取 (如 MZ、MA01，按 Enter 直接勾選)...`, showing `符合前綴共 M 個組別`, with `+ 勾選符合前綴 (M)` and `- 取消符合前綴` buttons.
  - Selected group chips row: Badges for each selected `groupCode` (with Chinese name) and `×` button to remove individual groups.
  - Scrollable multi-select checkbox grid of `filteredGroupOptions` so the user can also click individual groups to check/uncheck.

- [ ] **Step 4: Run test to verify it passes**

Run: `bun run typecheck && bun test && bun run build`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/App.tsx src/__tests__/app.test.ts docs/superpowers/specs/2026-10-31-group-match-rules-design.md docs/superpowers/plans/2026-10-31-prefix-multi-group-rules.md
git commit -m "feat(ui): support multi-select groups and prefix batch selection in group match rules"
```
