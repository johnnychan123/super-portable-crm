# Modal Group & Custom Composite Pickers Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move the "Target Groups Selection (Prefix + Multi-Select)" and "Custom Composite Column Combination" interfaces out of the main panel into dedicated Modal dialogs with draft state and "確認套用" / "取消" actions, keeping the main UI compact.

**Architecture:** Add compact summary formatters (`formatGroupSelectionSummary`, `formatCustomCompositeSummary`) and two modal draft state controllers (`groupPickerModal`, `customCompositeModal`) in `src/App.tsx`. On the main UI, render single-line summaries and trigger buttons (`選取組別...`, `設定複合欄...`), while all multi-select grids, Prefix batch controls, and column pickers live inside fixed overlay Modals that commit changes only on "確認套用".

**Tech Stack:** TypeScript, React 19, Tailwind CSS v4, Lucide React icons, Bun Test (`bun test`)

**Spec:** `docs/superpowers/specs/2026-10-31-group-match-rules-design.md`

## Global Constraints

- Do not render the multi-group checkbox grid, Prefix batch bar, or custom composite column checklist inline on the main panel; they must only appear inside their respective Modals.
- Canceling either Modal ("取消" or backdrop click) must leave the existing `defaultExternalMatchKey`, `groupMatchOverrides`, and `importedFilterFiles` untouched.
- The Custom Composite Modal's "確認套用" button must require at least 1 selected column (`draftColumns.length > 0`).

## Review Focus

- Selecting "自訂複合欄..." from a `<select>` dropdown opens the Custom Composite Modal with initialized draft columns without mutating the underlying `matchKey` until the user clicks "確認套用" (so clicking "取消" reverts the dropdown back to its prior value).
- Deleting a group rule or loading a new workbook resets any open modal states cleanly.
- `formatGroupSelectionSummary` handles 0 groups, `<= maxPreview` groups, and `> maxPreview` groups cleanly.
- `formatCustomCompositeSummary` handles empty `'custom:'` keys and multi-column `'custom:組別\u001f姓名'` keys cleanly.
- Pressing `Enter` in the Group Picker Modal's Prefix input batch-selects matching groups in `draftGroupCodes` without prematurely closing the modal.

---

### Task 1: Modal Pickers for Group Selection & Custom Composite Columns in `src/App.tsx`

**Files:**
- Modify: `src/App.tsx`
- Test: `src/__tests__/app.test.ts`

**Interfaces:**
- Produces in `src/App.tsx`:
  - `export function formatGroupSelectionSummary(groupCodes: ReadonlyArray<string>, maxPreview?: number): string`
  - `export function formatCustomCompositeSummary(matchKey: string): string`

- [ ] **Step 1: Write the failing test in `src/__tests__/app.test.ts`**

```typescript
    // Compact summary formatters for Modal trigger badges
    expect(formatGroupSelectionSummary([])).toBe('尚未選取組別');
    expect(formatGroupSelectionSummary(['MZ001', 'MZ002'], 3)).toBe(
      '已選 2 個組別（MZ001, MZ002）',
    );
    expect(formatGroupSelectionSummary(['MZ001', 'MZ002', 'MZ003', 'MZ004'], 3)).toBe(
      '已選 4 個組別（MZ001, MZ002, MZ003 等）',
    );
    expect(formatCustomCompositeSummary('custom:')).toBe('尚未選取欄位');
    expect(
      formatCustomCompositeSummary(encodeCustomMatchKey(['組別', '姓名', '手提電話'])),
    ).toBe('組別 + 姓名 + 手提電話');
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun test src/__tests__/app.test.ts`
Expected: FAIL with missing exports `formatGroupSelectionSummary` and `formatCustomCompositeSummary`

- [ ] **Step 3: Implement summary helpers and replace inline pickers with dedicated Modals in `src/App.tsx`**

1. Export `formatGroupSelectionSummary` and `formatCustomCompositeSummary` in `src/App.tsx`.
2. Replace `rulePrefixInputs` with two modal state objects:
   - `groupPickerModal: { ruleId: string; ruleIndex: number; draftGroupCodes: string[]; prefixInput: string } | null`
   - `customCompositeModal: { target: { type: 'default' } | { type: 'rule'; ruleId: string; ruleIndex: number } | { type: 'file'; fileId: string; fileName: string }; availableColumns: string[]; draftColumns: string[]; searchQuery: string } | null`
3. In `showExternalFilterPanel`:
   - For Default Match Key, Group Rule Match Key, and Single-File Match Key: when the user selects `CUSTOM_MATCH_KEY_PREFIX` from `<select>`, open `customCompositeModal` with `draftColumns` initialized from `parseCustomMatchKeyColumns(resolveInitialCustomMatchKey(currentKey, availableColumns))` instead of expanding an inline box. When the active `matchKey` is already a custom composite key, render a compact badge (`formatCustomCompositeSummary(matchKey)`) and a `"設定複合欄..."` button beside the `<select>`.
   - For each rule in `groupMatchOverrides`: render a compact single-row bar with `規則 #N`, `formatGroupSelectionSummary(selectedCodes)`, `"選取組別..."` button, the match key `<select>` (plus custom composite badge/button if applicable), and the delete button.
   - When `handleAddGroupMatchOverride` creates a new rule, automatically open `groupPickerModal` for that new rule.
4. Render the `groupPickerModal` and `customCompositeModal` overlay dialogs with draft state manipulation, backdrop close, `"取消"`, and `"確認套用"` buttons.

- [ ] **Step 4: Run typecheck, tests, and build to verify**

Run: `bun run typecheck && bun test && bun run build`
Expected: PASS (35 pass, 0 fail)

- [ ] **Step 5: Commit**

```bash
git add src/App.tsx src/__tests__/app.test.ts
git commit -m "feat: move group selection and custom composite column pickers into modals"
```
