import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, test } from 'bun:test';
import pkg from '../../package.json';
import { loadWindowContent, resolveAppPaths } from '../../electron/windowLoader';
import {
  APP_VERSION,
  applyBatchSelection,
  formatCustomCompositeSummary,
  formatGroupSelectionSummary,
  formatPlatformLabel,
  getGreetingMessage,
  getNextSortState,
  recomputeFilesWithGlobalRules,
  reconcileExternalFilterState,
  resolveActiveOutputHeaders,
  resolveInitialCustomMatchKey,
  toggleGroupCodeInList,
} from '../App';
import { encodeCustomMatchKey, parseCustomMatchKeyColumns } from '../lib/externalFilterMatcher';
import { computeVirtualWindow, findJumpTargetIndex } from '../components/VirtualizedDbTable';
import type { DbRowItem } from '../lib/excelParser';
import {
  matchExternalFilterToDb,
  type ImportedFilterFile,
  type ParsedExternalFilterSheet,
} from '../lib/externalFilterMatcher';

describe('Super Portable CRM Hello World', () => {
  test('getGreetingMessage returns expected Hello World greeting', () => {
    expect(getGreetingMessage()).toBe('Hello World! Welcome to Super Portable CRM');
    expect(getGreetingMessage('Custom CRM')).toBe('Hello World! Welcome to Custom CRM');
  });

  test('formatPlatformLabel maps OS platform identifiers accurately', () => {
    expect(formatPlatformLabel('win32', 'x64')).toBe('Windows (x64)');
    expect(formatPlatformLabel('darwin', 'arm64')).toBe('macOS (arm64)');
    expect(formatPlatformLabel('linux', 'x64')).toBe('Linux (x64)');
    expect(formatPlatformLabel('linux')).toBe('Linux');
  });

  test('APP_VERSION matches package.json version instead of stale 0.1.0', () => {
    expect(pkg.version).not.toBe('0.1.0');
    expect(APP_VERSION).toBe(pkg.version);
  });

  test('dev:electron script builds both renderer and electron before launching', () => {
    expect(pkg.scripts['dev:electron']).toContain('bun run build');
  });

  test('loadWindowContent falls back to loadFile when devServerUrl is unreachable', async () => {
    const calls: string[] = [];
    const target = {
      clearCache: async () => {
        calls.push('clearCache');
      },
      loadURL: async (_url: string) => {
        calls.push(`loadURL:${_url}`);
        throw new Error('ERR_CONNECTION_REFUSED');
      },
      loadFile: async (filePath: string) => {
        calls.push(`loadFile:${filePath}`);
      }
    };

    await loadWindowContent(target, {
      devServerUrl: 'http://localhost:5173',
      indexHtmlPath: '/app/dist/index.html'
    });

    expect(calls).toEqual([
      'clearCache',
      'loadURL:http://localhost:5173',
      'loadFile:/app/dist/index.html'
    ]);
  });

  test('electron-builder.yml disables win.signAndEditExecutable so dist:win works on Linux/WSL without wine', () => {
    const builderConfig = fs.readFileSync(
      path.join(__dirname, '../../electron-builder.yml'),
      'utf8'
    );
    expect(builderConfig).toContain('signAndEditExecutable: false');
  });

  test('ci.yml deploys dist/ to GitHub Pages and README.md documents the live web URL', () => {
    const ciWorkflow = fs.readFileSync(
      path.join(__dirname, '../../.github/workflows/ci.yml'),
      'utf8'
    );
    expect(ciWorkflow).toContain('actions/upload-pages-artifact');
    expect(ciWorkflow).toContain('actions/deploy-pages');

    const readme = fs.readFileSync(
      path.join(__dirname, '../../README.md'),
      'utf8'
    );
    expect(readme).toContain('https://johnnychan123.github.io/super-portable-crm/');
  });

  test('resolveAppPaths and bundled main.cjs resolve runtime paths via app.getAppPath without hardcoded __dirname', () => {
    const resolved = resolveAppPaths('/runtime/app.asar');
    expect(resolved.preloadPath).toBe(path.join('/runtime/app.asar', 'dist-electron', 'preload.cjs'));
    expect(resolved.indexHtmlPath).toBe(path.join('/runtime/app.asar', 'dist', 'index.html'));

    const mainSource = fs.readFileSync(
      path.join(__dirname, '../../electron/main.ts'),
      'utf8'
    );
    expect(mainSource).not.toContain('__dirname');
    expect(mainSource).toContain('app.getAppPath()');
    expect(mainSource).toContain('excel:save-file');
  });

  test('computeVirtualWindow calculates safe start/end indices and spacer heights for 0, small, and 10,000+ rows', () => {
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

    // Scrolled past bottom boundary safety check (clamps scrollTop to maxScrollTop so visible slice is non-empty)
    const bottomWin = computeVirtualWindow(999999, 600, 40, 12000, 10);
    expect(bottomWin.endIndex).toBe(12000);
    expect(bottomWin.bottomSpacerHeight).toBe(0);
    expect(bottomWin.startIndex).toBeLessThan(bottomWin.endIndex);
    expect(bottomWin.startIndex).toBe(11975);

    const mockRows = [
      { rowNumber: 2, kind: 'section' as const, sectionCode: 'MA000', values: {} },
      { rowNumber: 5, kind: 'record' as const, values: { 組別: 'MA001-01' } },
      { rowNumber: 49, kind: 'section' as const, sectionCode: 'MZ001', values: {} },
      { rowNumber: 50, kind: 'record' as const, values: { 組別: 'MZ001' } },
    ];
    expect(findJumpTargetIndex(mockRows, '#49')).toBe(2);
    expect(findJumpTargetIndex(mockRows, '50')).toBe(3);
    expect(findJumpTargetIndex(mockRows, 'mz001')).toBe(2);
    expect(findJumpTargetIndex(mockRows, 'notfound')).toBe(-1);
  });

  test('applyBatchSelection and getNextSortState handle selection operations and 3-state column sort cycling', () => {
    const initial = new Set<number>([5, 6]);
    const target = [6, 7, 8];

    expect(Array.from(applyBatchSelection(initial, target, 'select'))).toEqual([5, 6, 7, 8]);
    expect(Array.from(applyBatchSelection(initial, target, 'deselect'))).toEqual([5]);
    expect(Array.from(applyBatchSelection(initial, target, 'invert'))).toEqual([5, 7, 8]);
    expect(Array.from(applyBatchSelection(initial, target, 'clear'))).toEqual([]);

    const s0 = { column: null, direction: 'asc' as const };
    const s1 = getNextSortState(s0, '姓名');
    expect(s1).toEqual({ column: '姓名', direction: 'asc' });
    const s2 = getNextSortState(s1, '姓名');
    expect(s2).toEqual({ column: '姓名', direction: 'desc' });
    const s3 = getNextSortState(s2, '姓名');
    expect(s3).toEqual({ column: null, direction: 'asc' });
  });

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

  test('reconcileExternalFilterState cleans up stale exportTemplateFileId and intersection mode on file removal or clear', () => {
    // When 1 file remains after being in intersection mode with a removed template file ID:
    expect(
      reconcileExternalFilterState(['f2'], {
        externalFilterMode: 'intersection',
        exportTemplateFileId: 'f1',
        exportColumnMode: 'from_file',
      }),
    ).toEqual({
      externalFilterMode: 'all',
      exportTemplateFileId: 'f2',
      exportColumnMode: 'from_file',
    });

    // When all files are cleared:
    expect(
      reconcileExternalFilterState([], {
        externalFilterMode: 'union',
        exportTemplateFileId: 'f2',
        exportColumnMode: 'from_file',
      }),
    ).toEqual({
      externalFilterMode: 'all',
      exportTemplateFileId: '',
      exportColumnMode: 'db_all',
    });
  });

  test('recomputeFilesWithGlobalRules updates files using global mode while preserving single-file overrides', () => {
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

    // Multi-group groupCodes and toggleGroupCodeInList
    expect(toggleGroupCodeInList(['MZ001'], 'MZ002')).toEqual(['MZ001', 'MZ002']);
    expect(toggleGroupCodeInList(['MZ001', 'MZ002'], 'mz001')).toEqual(['MZ002']);

    // Custom composite match key initialization and global recomputation
    const initCustom = resolveInitialCustomMatchKey('auto', ['組別', '姓名', '手提電話']);
    expect(parseCustomMatchKeyColumns(initCustom)).toEqual(['組別', '姓名']);
    const existingCustom = encodeCustomMatchKey(['組別', '手提電話']);
    expect(resolveInitialCustomMatchKey(existingCustom, ['組別', '姓名', '手提電話'])).toBe(
      existingCustom,
    );

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
  });
});
