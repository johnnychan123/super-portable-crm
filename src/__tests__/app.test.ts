import { describe, expect, test } from 'bun:test';
import pkg from '../../package.json';
import { loadWindowContent } from '../../electron/windowLoader';
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

  test('APP_VERSION matches package.json version 0.2.1 instead of stale 0.1.0', () => {
    expect(pkg.version).toBe('0.2.1');
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
});
