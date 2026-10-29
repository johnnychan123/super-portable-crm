import React, { useEffect, useState } from 'react';
import type { AppRuntimeInfo } from './vite-env';

export function formatPlatformLabel(platform: string, arch?: string): string {
  const names: Record<string, string> = {
    win32: 'Windows',
    darwin: 'macOS',
    linux: 'Linux',
  };
  const label = names[platform] ?? platform;
  return arch ? `${label} (${arch})` : label;
}

export function getGreetingMessage(appName = 'Super Portable CRM'): string {
  return `Hello World! Welcome to ${appName}`;
}

export default function App(): React.ReactElement {
  const [info, setInfo] = useState<AppRuntimeInfo | null>(null);

  useEffect(() => {
    if (window.electronAPI?.getAppInfo) {
      window.electronAPI
        .getAppInfo()
        .then(setInfo)
        .catch(() => {
          // Fallback to static preload values if IPC is unavailable
        });
    }
  }, []);

  const platform = info?.platform ?? window.electronAPI?.platform ?? 'web';
  const arch = info?.arch ?? window.electronAPI?.arch ?? '';
  const electronVer = info?.electron ?? window.electronAPI?.versions.electron ?? 'N/A';
  const nodeVer = info?.node ?? window.electronAPI?.versions.node ?? 'N/A';

  return (
    <main className="app-container">
      <section className="hero-card">
        <span className="badge">Portable Desktop Edition</span>
        <h1>{getGreetingMessage()}</h1>
        <p className="subtitle">
          跨平台免安裝桌面應用程式 — 準備好連接您的 Excel 檔案進行資料處理與分析。
        </p>

        <div className="info-grid">
          <div className="info-item">
            <span className="info-label">作業系統</span>
            <strong className="info-value">{formatPlatformLabel(platform, arch)}</strong>
          </div>
          <div className="info-item">
            <span className="info-label">應用版本</span>
            <strong className="info-value">v{info?.version ?? '0.1.0'}</strong>
          </div>
          <div className="info-item">
            <span className="info-label">Electron</span>
            <strong className="info-value">{electronVer}</strong>
          </div>
          <div className="info-item">
            <span className="info-label">Node.js Runtime</span>
            <strong className="info-value">{nodeVer}</strong>
          </div>
        </div>

        <div className="placeholder-box">
          <h2>下一步：Excel 資料連接模組</h2>
          <p>基礎桌面環境與跨平台 CI/CD 管道已就緒，可隨時擴充 Excel 讀取與分析功能。</p>
        </div>
      </section>
    </main>
  );
}
