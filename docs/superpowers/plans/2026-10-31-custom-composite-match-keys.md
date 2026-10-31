# Custom Composite Match Keys Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Allow users to configure custom composite column combinations (e.g., `組別 + 姓名 + 手提電話` or `公司名稱 + 姓名`) for the Default Match Key, per-group override rules, and single-file match key overrides.

**Architecture:** Encode custom composite column selections as `custom:<col1>\u001f<col2>...` strings inside `matchKey` so `defaultExternalMatchKey`, `GroupMatchOverride.matchKey`, and `ImportedFilterFile.selectedMatchKey` share one uniform representation. Add `isCustomMatchKey`, `encodeCustomMatchKey`, `parseCustomMatchKeyColumns`, `toggleColumnInCustomMatchKey`, and `prioritizeMatchColumns` in `src/lib/externalFilterMatcher.ts`, and render an inline multi-column picker in `src/App.tsx` whenever a custom composite key is selected.

**Tech Stack:** TypeScript, React 19, Tailwind CSS v4, Bun Test (`bun test`)

**Spec:** `docs/superpowers/specs/2026-10-31-group-match-rules-design.md`

## Global Constraints

- When any column in a custom composite match key is missing from `sheet.sharedHeaders` or has a blank value on a record, that record must not match and must never fall back to partial columns.
- When `custom:` has 0 selected columns (`custom:`), `resolveMatchKeyColumns` returns `[]` and matches 0 rows safely.
- Keep `"dependencies": {}` empty in `package.json`.

## Review Focus

- Custom composite key with 3+ columns (`['組別', '姓名', '手提電話']`) only matches when all 3 columns are non-empty and equal (case-insensitive trimmed).
- Custom composite key where one of the selected columns is absent from the imported filter Excel returns `resolvedKeyColumns: []` and 0 matches instead of matching on the remaining columns.
- Empty `custom:` key (user switched dropdown to "自訂複合欄..." and cleared all columns) produces 0 matches without throwing or matching blanks.
- `toggleColumnInCustomMatchKey` preserves column order, avoids duplicates, and cleanly encodes/decodes column names containing `+` or parentheses (like `中文名(沒有稱謂)`).
- Switching a `<select>` dropdown to `'custom:'` initializes with sensible default columns (`['組別', '姓名']` if available) or preserves existing custom columns when already in custom mode.

---

### Task 1: Custom Composite Match Key Encoding & Matcher Resolution

**Files:**
- Modify: `src/lib/externalFilterMatcher.ts`
- Test: `src/__tests__/externalFilterMatcher.test.ts`

**Interfaces:**
- Produces in `src/lib/externalFilterMatcher.ts`:
  - `export const CUSTOM_MATCH_KEY_PREFIX = 'custom:';`
  - `export function isCustomMatchKey(matchKey: string): boolean`
  - `export function encodeCustomMatchKey(columns: ReadonlyArray<string>): string`
  - `export function parseCustomMatchKeyColumns(matchKey: string): string[]`
  - `export function toggleColumnInCustomMatchKey(matchKey: string, column: string): string`
  - `export function prioritizeMatchColumns(headers: ReadonlyArray<string>): string[]`
  - Updated `resolveMatchKeyColumns(matchKey, sheet, dbRows)` to support `isCustomMatchKey(matchKey)`
  - Updated `getGlobalMatchKeyOptions(dbHeaders)` and `getAvailableMatchKeys(sheet, dbRows)` to include `{ key: 'custom:', label: '自訂複合欄（任意勾選多欄組合）...', columns: [] }`

- [ ] **Step 1: Write the failing test in `src/__tests__/externalFilterMatcher.test.ts`**

```typescript
  it('supports custom composite match keys (2+ arbitrary columns) in default, group override, and single-file modes', () => {
    const dbHeaders = ['組別', 'PID', '姓名', '中文名(沒有稱謂)', '手提電話', '公司名稱'];
    const dbRows: DbRowItem[] = [
      makeRecordRow(2, 'MA001', {
        組別: 'MA001',
        PID: '1',
        姓名: '陳大文',
        '中文名(沒有稱謂)': '陳大文',
        手提電話: '91112222',
        公司名稱: 'Alpha Ltd',
      }),
      makeRecordRow(3, 'MA001', {
        組別: 'MA001',
        PID: '2',
        姓名: '陳大文',
        '中文名(沒有稱謂)': '陳大文',
        手提電話: '93334444',
        公司名稱: 'Beta Ltd',
      }),
      makeRecordRow(4, 'MZ001', {
        組別: 'MZ001',
        PID: '',
        姓名: '李小明',
        '中文名(沒有稱謂)': '李小明',
        手提電話: '',
        公司名稱: 'Gamma Corp',
      }),
    ];

    const extBuf = buildWorkbookBuffer([
      ['組別', '姓名', '中文名(沒有稱謂)', '手提電話', '公司名稱'],
      ['MA001', '陳大文', '陳大文', '93334444', 'Beta Ltd'],
      ['MZ001', '李小明', '李小明', '', 'Gamma Corp'],
    ]);
    const sheet = parseExternalFilterWorkbook(extBuf, 'custom-composite.xlsx', dbHeaders);

    // 1. Encode/decode and toggle helpers
    const key1 = encodeCustomMatchKey(['組別', '姓名', '手提電話']);
    expect(isCustomMatchKey(key1)).toBe(true);
    expect(parseCustomMatchKeyColumns(key1)).toEqual(['組別', '姓名', '手提電話']);

    const toggled = toggleColumnInCustomMatchKey(key1, '姓名');
    expect(parseCustomMatchKeyColumns(toggled)).toEqual(['組別', '手提電話']);
    const toggledBack = toggleColumnInCustomMatchKey(toggled, '中文名(沒有稱謂)');
    expect(parseCustomMatchKeyColumns(toggledBack)).toEqual(['組別', '手提電話', '中文名(沒有稱謂)']);

    // 2. Prioritized columns put common identity columns first
    const prioritized = prioritizeMatchColumns(['備註', '公司名稱', '組別', '手提電話', '姓名']);
    expect(prioritized.slice(0, 4)).toEqual(['組別', '姓名', '手提電話', '公司名稱']);

    // 3. Match with 3-column custom key ['組別', '姓名', '手提電話'] disambiguates same-name records in MA001
    // and strictly does not match MZ001 where 手提電話 is blank
    const res3Col = matchExternalFilterToDb(sheet, dbRows, key1);
    expect(res3Col.resolvedKeyColumns).toEqual(['組別', '姓名', '手提電話']);
    expect(Array.from(res3Col.matchedDbRowNumbers)).toEqual([3]);

    // 4. Global default = key1 (3-col), plus MZ001 group override using custom ['組別', '公司名稱', '中文名(沒有稱謂)']
    const mzCustomKey = encodeCustomMatchKey(['組別', '公司名稱', '中文名(沒有稱謂)']);
    const resGlobal = matchExternalFilterToDb(sheet, dbRows, 'global', {
      defaultMatchKey: key1,
      groupOverrides: [
        { id: 'r1', groupCodes: ['MZ001'], matchKey: mzCustomKey },
      ],
    });
    expect(Array.from(resGlobal.matchedDbRowNumbers).sort((a, b) => a - b)).toEqual([3, 4]);

    // 5. Missing column in external sheet or empty custom key returns 0 matches safely
    const missingColKey = encodeCustomMatchKey(['組別', 'PID']); // PID not in extBuf
    expect(Array.from(matchExternalFilterToDb(sheet, dbRows, missingColKey).matchedDbRowNumbers)).toEqual([]);
    expect(Array.from(matchExternalFilterToDb(sheet, dbRows, 'custom:').matchedDbRowNumbers)).toEqual([]);
  });
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun test src/__tests__/externalFilterMatcher.test.ts`
Expected: FAIL with missing exports `encodeCustomMatchKey`, `isCustomMatchKey`, etc.

- [ ] **Step 3: Implement custom composite match key helpers and resolution in `src/lib/externalFilterMatcher.ts`**

Add `CUSTOM_MATCH_KEY_PREFIX = 'custom:'`, `CUSTOM_COL_DELIMITER = '\u001f'`, `isCustomMatchKey`, `encodeCustomMatchKey`, `parseCustomMatchKeyColumns`, `toggleColumnInCustomMatchKey`, and `prioritizeMatchColumns`, and update `resolveMatchKeyColumns`, `getGlobalMatchKeyOptions`, and `getAvailableMatchKeys`.

- [ ] **Step 4: Run test to verify it passes**

Run: `bun test src/__tests__/externalFilterMatcher.test.ts`
Expected: PASS (7 pass, 0 fail)

- [ ] **Step 5: Commit**

```bash
git add src/lib/externalFilterMatcher.ts src/__tests__/externalFilterMatcher.test.ts
git commit -m "feat: support custom composite match keys in externalFilterMatcher"
```

---

### Task 2: Inline Custom Composite Column Selector UI in `src/App.tsx`

**Files:**
- Modify: `src/App.tsx`
- Test: `src/__tests__/app.test.ts`

**Interfaces:**
- Consumes from `src/lib/externalFilterMatcher.ts`:
  - `isCustomMatchKey`, `encodeCustomMatchKey`, `parseCustomMatchKeyColumns`, `toggleColumnInCustomMatchKey`, `prioritizeMatchColumns`
- Produces in `src/App.tsx`:
  - `export function resolveInitialCustomMatchKey(currentKey: string, availableColumns: ReadonlyArray<string>): string`
  - Inline custom column selector UI in `showExternalFilterPanel` for:
    1. Global Default Match Key (`defaultExternalMatchKey`)
    2. Each Group Match Override rule (`rule.matchKey`)
    3. Each Single-File Match Key override (`file.selectedMatchKey`)

- [ ] **Step 1: Write the failing test in `src/__tests__/app.test.ts`**

Add a test verifying `resolveInitialCustomMatchKey` and `recomputeFilesWithGlobalRules` with custom composite match keys (`encodeCustomMatchKey(['組別', '公司名稱'])`).

- [ ] **Step 2: Run test to verify it fails**

Run: `bun test src/__tests__/app.test.ts`
Expected: FAIL with missing export `resolveInitialCustomMatchKey`

- [ ] **Step 3: Implement `resolveInitialCustomMatchKey` and inline custom composite column picker UI in `src/App.tsx`**

1. Export `resolveInitialCustomMatchKey(currentKey: string, availableColumns: ReadonlyArray<string>): string`:
   - If `isCustomMatchKey(currentKey)` already has columns, return `currentKey`.
   - Otherwise, pick `['組別', '姓名']` if both exist in `availableColumns`, else the first 2 columns of `prioritizeMatchColumns(availableColumns)`, encoded via `encodeCustomMatchKey`.
2. In `src/App.tsx`, render a reusable inline custom composite column picker whenever `isCustomMatchKey(...)` is active on:
   - Default Match Key (`defaultExternalMatchKey`)
   - Group Override Rule (`rule.matchKey`)
   - Single-File Match Key (`file.selectedMatchKey`)
   Showing ordered selected column badges (`1. 組別 + 2. 姓名 + ...` with `×` remove and "清空") and clickable column toggle pills (`prioritizeMatchColumns`).

- [ ] **Step 4: Run typecheck, tests, and build to verify**

Run: `bun run typecheck && bun test && bun run build`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/App.tsx src/__tests__/app.test.ts
git commit -m "feat: add inline custom composite column selector UI for match keys"
```
