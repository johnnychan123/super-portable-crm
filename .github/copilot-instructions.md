# Copilot Instructions — Super Portable CRM

## Build, Test, and Typecheck Commands

This project uses **Bun** for package management, testing, and Electron main/preload bundling, and **Vite** for the React renderer.

```bash
# Install dependencies
bun install

# Typecheck (TypeScript across src/, electron/, and vite.config.ts)
bun run typecheck

# Run all unit tests
bun test

# Run a single test file or filter by test name
bun test src/__tests__/app.test.ts
bun test src/__tests__/excelParser.test.ts
bun test -t "parseCrmWorkbook"

# Development
# Option A: Run Vite dev server for HMR, then launch Electron in another terminal
bun run dev:renderer
bun run dev:electron
# Option B: Run `bun run dev:electron` standalone (builds first and falls back to dist/index.html if :5173 is unreachable)

# Build renderer (dist/) and Electron main/preload (dist-electron/)
bun run build

# Package portable binaries into release/ (use CSC_IDENTITY_AUTO_DISCOVERY=false when unsigned)
bun run dist:win
bun run dist:mac
bun run dist:linux

# Run packaged Linux binary on Ubuntu 22.04+/WSL2 without libfuse2
./release/linux-unpacked/super-portable-crm
# or: ./release/Super-Portable-CRM-*-Linux-x86_64.AppImage --appimage-extract-and-run
```

## High-Level Architecture

- **Dual-Build Pipeline**:
  - **Renderer (`src/` -> `dist/`)**: Built with Vite + React 19 + Tailwind CSS v4 (`@tailwindcss/vite`). `vite.config.ts` sets `base: './'` so `mainWindow.loadFile` resolves relative asset paths inside `app.asar`.
  - **Electron Main & Preload (`electron/` -> `dist-electron/`)**: Bundled via `bun build` into CommonJS (`main.cjs` and `preload.cjs`) with `--target=node --format=cjs --external=electron`.
  - **Window Loading (`electron/windowLoader.ts`)**: `resolveAppPaths(app.getAppPath())` computes runtime paths for `dist-electron/preload.cjs` and `dist/index.html`. When `VITE_DEV_SERVER_URL` is set, `loadWindowContent` clears the session cache and attempts `loadURL`, falling back to `loadFile(indexHtmlPath)` if the dev server is not running.
- **IPC & Dual-Mode File Loading**:
  - `electron/main.ts` creates the `BrowserWindow` with `contextIsolation: true` and `nodeIntegration: false`, registering IPC handlers via `ipcMain.handle(...)` (`app:get-info`, `excel:open-file`).
  - `electron/preload.ts` bridges main-process capabilities to the renderer via `contextBridge.exposeInMainWorld('electronAPI', ...)`.
  - `src/App.tsx` supports both Electron native file dialogs (`window.electronAPI.openExcelFile()`) and browser `<input type="file">` / drag-and-drop fallback (`File.arrayBuffer()`), feeding a `Uint8Array` into `parseCrmWorkbook`.
  - Any new IPC method added in `electron/main.ts` and `electron/preload.ts` must also be typed on `ElectronAPI` (`Window['electronAPI']`) in `src/vite-env.d.ts`.
- **CRM Workbook Parser & Domain Model (`src/lib/excelParser.ts`)**:
  - Uses SheetJS (`xlsx`) to parse multi-sheet `.xlsx`/`.xls` workbooks containing `DB` (or first sheet fallback), `組別` (groups), and `界別` (sectors).
  - Normalizes Excel carriage-return artifacts (`_x000D_`) and classifies `DB` rows into `'section'` header rows (matching `/^(M[A-Z]\d{3}(?:-\d{2})?)\s*(.*)$/` when `組別`, `計算`, and `出席活動` are blank) vs `'record'` rows.
  - Propagates `currentSectionLabel` onto following records, uses discovered section titles to backfill masked (`XXXX`) or empty Chinese names in `組別`/`界別`, and provides column presets (`core`, `contact`, `personal`, `all` via `getVisibleDbHeaders`) and filtering helpers (`filterDbRows`, `filterGroups`).

## Key Conventions & Packaging Rules

- **Bundle Dependencies at Build Time**: Keep `"dependencies": {}` empty in `package.json` and install libraries in `"devDependencies"` whenever they can be bundled by Vite (`dist/`) or `bun build` (`dist-electron/`). `electron-builder.yml` is configured with `npmRebuild: false` and packages only `dist/**/*`, `dist-electron/**/*`, and `package.json`.
- **`bun build` Flag Syntax & Runtime Paths**:
  - Always pass `--external=electron` (with `=`) in `build:electron` scripts so `bun build` does not mistake `electron` for the local `./electron/` directory.
  - Never use `__dirname` or `__filename` inside `electron/` source files because `bun build` inlines them at compile time as hardcoded build-machine paths. Always resolve `dist/index.html` and `dist-electron/preload.cjs` at runtime via `app.getAppPath()` (`resolveAppPaths(app.getAppPath())`).
- **Version Sync**: `src/App.tsx` exports `APP_VERSION = pkg.version` directly from `package.json`. `.github/workflows/release.yml` automatically syncs `package.json`'s `"version"` field from the release tag (`vX.Y.Z`) prior to running tests and `electron-builder`.
- **WSL2 / Linux Windows Cross-Packaging**: `electron-builder.yml` sets `win.signAndEditExecutable: false` so `bun run dist:win` can package the Windows NSIS portable `.exe` on Linux/WSL2 without requiring Wine.
- **Git & Release Workflow**:
  - In environments where direct `git push` is blocked by proxy, push branch commits using `git push-api`.
  - Trigger a release either via `gh workflow run release.yml -f tag=vX.Y.Z` or by creating a `v*` tag ref via `gh api -X POST repos/johnnychan123/super-portable-crm/git/refs -f ref="refs/tags/vX.Y.Z" -f sha="$(git rev-parse HEAD)"`.
  - `.github/workflows/release.yml` builds Windows (`portable` `.exe`), macOS (`zip` & `dmg`), and Linux (`AppImage`) artifacts and attaches them to the GitHub Release.
