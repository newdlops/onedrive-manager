# OneDrive Manager

Cross-platform Electron desktop app for managing Microsoft OneDrive on Windows and macOS.

## Platform Policy

- Keep OS-specific behavior in the Electron main process behind small platform modules.
- Use standard Node and Electron APIs before adding native dependencies.
- Treat Windows as a primary runtime target even when developing on macOS.
- Build and verify packages on native GitHub Actions runners for both Windows and macOS.
- Avoid renderer-side filesystem access. Use the preload bridge and typed IPC.

## Scripts

```sh
npm run dev
npm run typecheck
npm run build
npm run package
npm run dist:mac
npm run dist:win
```

`dist:win` is available from macOS for basic cross-builds, but release builds should run on Windows CI because installer tooling, signing, and filesystem behavior are platform-sensitive.

## Structure

```text
src/main/       Electron main process and platform-safe filesystem logic
src/preload/    Typed IPC bridge exposed to the renderer
src/renderer/   React UI
src/shared/     Types shared by main, preload, and renderer
```

