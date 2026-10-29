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

const electronAPI = {
  platform: process.platform,
  arch: process.arch,
  versions: {
    electron: process.versions.electron,
    node: process.versions.node,
    chrome: process.versions.chrome,
  },
  getAppInfo: (): Promise<AppRuntimeInfo> => ipcRenderer.invoke('app:get-info'),
};

contextBridge.exposeInMainWorld('electronAPI', electronAPI);
