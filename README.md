# Super Portable CRM

跨平台免安裝（Portable）CRM 與 Excel 資料處理桌面應用程式，採用 **Bun + Electron + React + TypeScript (Vite)** 架構。

## 環境需求

- [Bun](https://bun.sh/) >= 1.1

## 常用指令

```bash
# 1. 安裝依賴
bun install

# 2. 執行型別檢查與單元測試
bun run typecheck
bun test

# 3. 建置前端與 Electron 主程序
bun run build

# 4. 本地打包對應平台的 Portable 執行檔（輸出於 release/ 目錄）
bun run dist:win    # Windows Portable (.exe)
bun run dist:mac    # macOS (.zip / .dmg)
bun run dist:linux  # Linux (.AppImage)
```

## 觸發 GitHub Actions 自動打包與發佈 Release

本專案已配置 `.github/workflows/release.yml`，支援自動建置 **Windows / macOS / Linux** 三平台執行檔並上傳至 **GitHub Releases**：

### 方式一：透過 GitHub CLI 或網頁手動觸發（推薦，無須直接 git push tag）
```bash
# 先用 git push-api 推送程式碼到 main
git push-api

# 使用 gh CLI 觸發三平台打包並發佈 v0.1.0 Release
gh workflow run release.yml -f tag=v0.1.0
```
或前往 GitHub Repository 的 **Actions** -> **Release Portable Binaries** -> **Run workflow** 輸入版本號（如 `v0.1.0`）。

### 方式二：推送版本 Tag（一般網路環境）
```bash
git tag v0.1.0
git push origin v0.1.0
```

完成後即可在 **GitHub Releases** 下載：
- **Windows**：`Super-Portable-CRM-0.1.0-Windows-x64-Portable.exe`
- **macOS**：`Super-Portable-CRM-0.1.0-macOS-arm64.zip` / `.dmg` 及 `x64` 版本
- **Linux**：`Super-Portable-CRM-0.1.0-Linux-x64.AppImage`

