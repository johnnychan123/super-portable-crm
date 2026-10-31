import { app, BrowserWindow, dialog, ipcMain } from 'electron';
import fs from 'node:fs/promises';
import path from 'node:path';
import { loadWindowContent, resolveAppPaths } from './windowLoader';

function createWindow(): void {
  const { preloadPath, indexHtmlPath } = resolveAppPaths(app.getAppPath());

  const mainWindow = new BrowserWindow({
    width: 1280,
    height: 820,
    minWidth: 900,
    minHeight: 600,
    title: 'Super Portable CRM',
    webPreferences: {
      preload: preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  const devServerUrl = process.env.VITE_DEV_SERVER_URL;

  void loadWindowContent(
    {
      clearCache: () => mainWindow.webContents.session.clearCache(),
      loadURL: (url) => mainWindow.loadURL(url),
      loadFile: (filePath) => mainWindow.loadFile(filePath),
    },
    { devServerUrl, indexHtmlPath },
  );
}

app.whenReady().then(() => {
  ipcMain.handle('app:get-info', () => ({
    name: app.getName(),
    version: app.getVersion(),
    platform: process.platform,
    arch: process.arch,
    electron: process.versions.electron,
    node: process.versions.node,
    chrome: process.versions.chrome,
  }));

  ipcMain.handle('excel:open-file', async () => {
    const result = await dialog.showOpenDialog({
      title: '選取要匯入的 Excel 檔案',
      properties: ['openFile'],
      filters: [{ name: 'Excel 活頁簿 (*.xlsx, *.xls)', extensions: ['xlsx', 'xls'] }],
    });

    if (result.canceled || result.filePaths.length === 0) {
      return { canceled: true };
    }

    const filePath = result.filePaths[0];
    const fileBuffer = await fs.readFile(filePath);
    return {
      canceled: false,
      fileName: path.basename(filePath),
      filePath,
      buffer: new Uint8Array(fileBuffer),
    };
  });

  ipcMain.handle(
    'excel:save-file',
    async (_event, payload: { defaultFileName: string; buffer: Uint8Array }) => {
      const result = await dialog.showSaveDialog({
        title: '儲存匯出的 Excel 檔案',
        defaultPath: payload.defaultFileName,
        filters: [{ name: 'Excel 活頁簿 (*.xlsx)', extensions: ['xlsx'] }],
      });

      if (result.canceled || !result.filePath) {
        return { canceled: true };
      }

      await fs.writeFile(result.filePath, Buffer.from(payload.buffer));
      return {
        canceled: false,
        filePath: result.filePath,
      };
    },
  );

  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
