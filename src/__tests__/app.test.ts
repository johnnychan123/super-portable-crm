import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, test } from 'bun:test';
import pkg from '../../package.json';
import { loadWindowContent, resolveAppPaths } from '../../electron/windowLoader';
import { APP_VERSION, formatPlatformLabel, getGreetingMessage } from '../App';

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
  });
});
