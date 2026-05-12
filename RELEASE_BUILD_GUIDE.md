# Release Build Guide — Customer Recovery Console

This project is now configured to generate distributable installers:

- Windows: `.exe` installer using NSIS
- macOS: `.dmg` disk image, plus `.zip`

## Local prerequisites

Install Node.js 22 LTS or another Vite-compatible modern Node release.

From the project root:

```bash
npm install
npm run check
npm run build
```

## Build Windows installer

Recommended on a Windows machine or a Windows GitHub Actions runner:

```bash
npm run dist:win
```

Output appears in:

```text
release/
```

Expected artifact pattern:

```text
Customer Recovery Console-0.2.0-win-x64.exe
```

## Build macOS DMG

Recommended on a Mac or macOS GitHub Actions runner:

```bash
npm run dist:mac
```

Output appears in:

```text
release/
```

Expected artifact patterns:

```text
Customer Recovery Console-0.2.0-mac-x64.dmg
Customer Recovery Console-0.2.0-mac-arm64.dmg
```

## Build both from macOS only

macOS can often build both Mac and Windows installer artifacts for this Electron app:

```bash
npm run dist:all-local
```

For the cleanest production pipeline, use GitHub Actions with one Windows runner and one macOS runner.

## GitHub Actions release build

A workflow has been added at:

```text
.github/workflows/build-installers.yml
```

To use it:

1. Push this project to GitHub.
2. Open the repository in GitHub.
3. Go to Actions.
4. Select **Build desktop installers**.
5. Click **Run workflow**.
6. Download the generated artifacts:
   - Windows installer zip containing `.exe`
   - macOS DMG zip containing `.dmg`

## Signing and notarization status

The default local and GitHub builds are unsigned.

For internal testing, unsigned builds are usually acceptable.

For client-facing distribution, plan for:

- Windows code-signing certificate
- Apple Developer ID certificate
- Apple notarization for the macOS `.dmg`

Without signing/notarization, Windows and macOS may warn users before opening the installer/app.

## Recommended client delivery package

For a non-technical client, send a simple ZIP package:

```text
Customer-Recovery-Console-Windows.zip
  Customer Recovery Console-0.2.0-win-x64.exe
  READ_ME_FIRST_WINDOWS.pdf or .txt

Customer-Recovery-Console-Mac.zip
  Customer Recovery Console-0.2.0-mac-arm64.dmg
  Customer Recovery Console-0.2.0-mac-x64.dmg
  READ_ME_FIRST_MAC.pdf or .txt
```

## Which Mac build should the client use?

- Apple Silicon / M1 / M2 / M3 / M4 Mac: use `arm64.dmg`
- Older Intel Mac: use `x64.dmg`

If unsure, send both and tell the client to open **Apple menu → About This Mac**.
