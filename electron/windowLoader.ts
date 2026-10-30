import path from 'node:path';

export interface WindowContentTarget {
  clearCache: () => Promise<void>;
  loadURL: (url: string) => Promise<void>;
  loadFile: (filePath: string) => Promise<void>;
}

export interface LoadWindowContentOptions {
  devServerUrl?: string;
  indexHtmlPath: string;
}

export interface ResolvedAppPaths {
  preloadPath: string;
  indexHtmlPath: string;
}

export function resolveAppPaths(appPath: string): ResolvedAppPaths {
  return {
    preloadPath: path.join(appPath, 'dist-electron', 'preload.cjs'),
    indexHtmlPath: path.join(appPath, 'dist', 'index.html'),
  };
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
