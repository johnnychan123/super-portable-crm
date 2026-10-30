# Copilot Instructions — Super Portable CRM

## Build, Test, and Typecheck Commands

This project uses **Bun** for package management, testing, and Electron main/preload bundling, and **Vite** for the React renderer.

```bash
# Install dependencies
bun install

# Typecheck (TypeScript across src/ and electron/)
bun run typecheck

# Run all unit tests
bun test

# Run a single test file or filter by test name
bun test src/__tests__/app.test.ts
bun test -t "getGreetingMessage"

# Development (run Vite dev server first, then launch Electron in another terminal)
bun run dev:renderer
bun run dev:electron

# Build renderer (dist/) and Electron main/preload (dist-electron/)
bun run build

# Package portable binaries into release/ (use CSC_IDENTITY_AUTO_DISCOVERY=false when unsigned)
bun run dist:linux
bun run dist:win
bun run dist:mac
```

## High-Level Architecture

- **Dual-Build Pipeline**:
  - **Renderer (`src/` -> `dist/`)**: Built with Vite + React 19 (`vite.config.ts` sets `base: './'` so `mainWindow.loadFile` resolves relative asset paths inside `app.asar`).
  - **Electron Main & Preload (`electron/` -> `dist-electron/`)**: Bundled via `bun build` into CommonJS (`main.cjs` and `preload.cjs`) with `--target=node --format=cjs --external=electron`.
- **IPC & Security Boundary**:
  - `electron/main.ts` creates the `BrowserWindow` with `contextIsolation: true` and `nodeIntegration: false`, registering IPC handlers via `ipcMain.handle(...)`.
  - `electron/preload.ts` bridges main-process capabilities to the renderer via `contextBridge.exposeInMainWorld('electronAPI', ...)`.
  - Any new IPC method added in `electron/main.ts` and `electron/preload.ts` must also be typed on `Window['electronAPI']` in `src/vite-env.d.ts`.

## Key Conventions & Packaging Rules

- **Bundle Dependencies at Build Time**: Keep `"dependencies": {}` empty in `package.json` and install libraries in `"devDependencies"` whenever they can be bundled by Vite (`dist/`) or `bun build` (`dist-electron/`). `electron-builder.yml` is configured with `npmRebuild: false` and packages only `dist/**/*`, `dist-electron/**/*`, and `package.json`.
- **`bun build` Flag Syntax & Runtime Paths**:
  - Always pass `--external=electron` (with `=`) in `build:electron` scripts so `bun build` does not mistake `electron` for the local `./electron/` directory.
  - Never use `__dirname` or `__filename` inside `electron/` source files because `bun build` replaces them at compile time with the build machine's hardcoded source path. Always resolve `dist/index.html` and `dist-electron/preload.cjs` at runtime via `app.getAppPath()` (`resolveAppPaths(app.getAppPath())`).
- **WSL2 / Linux Windows Cross-Packaging**: `electron-builder.yml` sets `win.signAndEditExecutable: false` so `bun run dist:win` can package the Windows NSIS portable `.exe` on Linux/WSL2 without requiring Wine.
- **Git & Release Workflow**:
  - In environments where direct `git push` is blocked by proxy, push branch commits using `git push-api`.
  - Trigger a release either via `gh workflow run release.yml -f tag=vX.Y.Z` or by creating a `v*` tag ref via `gh api -X POST repos/johnnychan123/super-portable-crm/git/refs -f ref="refs/tags/vX.Y.Z" -f sha="$(git rev-parse HEAD)"`.
  - `.github/workflows/release.yml` builds Windows (`portable` `.exe`), macOS (`zip` & `dmg`), and Linux (`AppImage`) artifacts and attaches them to the GitHub Release.
