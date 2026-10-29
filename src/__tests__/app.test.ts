import { describe, expect, test } from 'bun:test';
import { formatPlatformLabel, getGreetingMessage } from '../App';

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
});
