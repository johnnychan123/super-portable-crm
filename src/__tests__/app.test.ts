import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, test } from 'bun:test';
import pkg from '../../package.json';
import { loadWindowContent, resolveAppPaths } from '../../electron/windowLoader';
import {
  APP_VERSION,
  applyBatchSelection,
  formatPlatformLabel,
  getGreetingMessage,
  getNextSortState,
} from '../App';
import { computeVirtualWindow } from '../components/VirtualizedDbTable';

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

    // Scrolled past bottom boundary safety check
    const bottomWin = computeVirtualWindow(999999, 600, 40, 12000, 10);
    expect(bottomWin.endIndex).toBe(12000);
    expect(bottomWin.bottomSpacerHeight).toBe(0);
    expect(bottomWin.startIndex).toBeLessThanOrEqual(12000);
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
});
