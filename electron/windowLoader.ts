export interface WindowContentTarget {
  clearCache: () => Promise<void>;
  loadURL: (url: string) => Promise<void>;
  loadFile: (filePath: string) => Promise<void>;
}

export interface LoadWindowContentOptions {
  devServerUrl?: string;
  indexHtmlPath: string;
}

export async function loadWindowContent(
  target: WindowContentTarget,
  options: LoadWindowContentOptions
): Promise<void> {
  const { devServerUrl, indexHtmlPath } = options;

  if (devServerUrl) {
    await target.clearCache();
    try {
      await target.loadURL(devServerUrl);
      return;
    } catch {
      await target.loadFile(indexHtmlPath);
      return;
    }
  }

  await target.loadFile(indexHtmlPath);
}
