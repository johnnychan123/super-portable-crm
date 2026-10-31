# Super Portable CRM

跨平台免安裝（Portable）CRM 與 Excel 資料處理桌面應用程式，採用 **Bun + Electron + React + TypeScript (Vite)** 架構。

## 線上網頁版（GitHub Pages）

無需下載或安裝，直接於瀏覽器開啟即可使用（所有 Excel 解析、比對與匯出皆在瀏覽器本地端執行，不會上傳至任何伺服器）：

- **線上網址**：[https://johnnychan123.github.io/super-portable-crm/](https://johnnychan123.github.io/super-portable-crm/)
- **自動部署**：每次推送至 `main` 分支時，`.github/workflows/ci.yml` 會自動執行測試、建置 `dist/` 並發佈至 GitHub Pages。

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

### 方式一：透過 GitHub CLI 觸發或建立版本 Tag（推薦，適用於本機代理環境）
```bash
# 1. 先用 git push-api 推送程式碼到 main
git push-api

# 2A. 直接用 gh workflow run 觸發三平台打包並發佈 Release
gh workflow run release.yml -f tag=v0.1.0

# 2B. 或透過 GitHub Git Refs API 建立遠端 Tag 自動觸發
gh api -X POST repos/johnnychan123/super-portable-crm/git/refs \
  -f ref="refs/tags/v0.1.0" \
  -f sha="$(git rev-parse HEAD)"
```

### 方式二：GitHub 網頁手動觸發
前往 GitHub Repository 的 **Actions** -> **Release Portable Binaries** -> **Run workflow** 輸入版本號（如 `v0.1.0`）。

### 方式三：推送版本 Tag（一般網路環境）
```bash
git tag v0.1.0
git push origin v0.1.0
```

完成後即可在 **GitHub Releases** 下載：
- **Windows**：`Super-Portable-CRM-0.1.0-Windows-x64-Portable.exe`
- **macOS**：`Super-Portable-CRM-0.1.0-macOS-arm64.zip` / `.dmg` 及 `x64` 版本
- **Linux**：`Super-Portable-CRM-0.1.0-Linux-x64.AppImage`

