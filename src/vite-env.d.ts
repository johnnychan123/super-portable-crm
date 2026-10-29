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

export interface ElectronAPI {
  platform: string;
  arch: string;
  versions: {
    electron: string;
    node: string;
    chrome: string;
  };
  getAppInfo: () => Promise<AppRuntimeInfo>;
}

declare global {
  interface Window {
    electronAPI?: ElectronAPI;
  }
}
