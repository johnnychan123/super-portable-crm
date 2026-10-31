import { contextBridge, ipcRenderer } from 'electron';

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

export interface ExcelSaveFilePayload {
  defaultFileName: string;
  buffer: Uint8Array;
}

export interface ExcelSaveFileResult {
  canceled: boolean;
  filePath?: string;
}

const electronAPI = {
  platform: process.platform,
  arch: process.arch,
  versions: {
    electron: process.versions.electron,
    node: process.versions.node,
    chrome: process.versions.chrome,
  },
  getAppInfo: (): Promise<AppRuntimeInfo> => ipcRenderer.invoke('app:get-info'),
  openExcelFile: (): Promise<ExcelOpenFileResult> => ipcRenderer.invoke('excel:open-file'),
  saveExcelFile: (payload: ExcelSaveFilePayload): Promise<ExcelSaveFileResult> =>
    ipcRenderer.invoke('excel:save-file', payload),
};

contextBridge.exposeInMainWorld('electronAPI', electronAPI);
