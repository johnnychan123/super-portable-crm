/// <reference types="vite/client" />

export interface AppRuntimeInfo {
  name: string;
  version: string;
  platform: string;
  arch: string;
  electron: string;
  node: string;
  chrome: string;
}

export interface ExcelOpenFileResult {
  canceled: boolean;
  fileName?: string;
  filePath?: string;
  buffer?: Uint8Array;
}

export interface ExcelOpenedFileItem {
  fileName: string;
  filePath?: string;
  buffer: Uint8Array;
}

export interface ExcelOpenMultipleFilesResult {
  canceled: boolean;
  files?: ExcelOpenedFileItem[];
}

export interface ExcelSaveFilePayload {
  defaultFileName: string;
  buffer: Uint8Array;
}

export interface ExcelSaveFileResult {
  canceled: boolean;
  filePath?: string;
}

export interface ElectronAPI {
  platform: string;
  arch: string;
  versions: {
    electron: string;
    node: string;
    chrome: string;
  };
  getAppInfo: () => Promise<AppRuntimeInfo>;
  openExcelFile?: () => Promise<ExcelOpenFileResult>;
  openMultipleExcelFiles?: () => Promise<ExcelOpenMultipleFilesResult>;
  saveExcelFile?: (payload: ExcelSaveFilePayload) => Promise<ExcelSaveFileResult>;
}

declare global {
  interface Window {
    electronAPI?: ElectronAPI;
  }
}
