# Customer Recovery Console — MVP v0.2

A local-first Electron + React app for recovering, cleaning, reviewing, and exporting customer records from authorized customer/export files.

## What this MVP does

- Runs as a local desktop app.
- Imports CSV, TXT, XLS, and XLSX files.
- Auto-detects email, name, order ID, date, amount, and consent-like columns.
- Deduplicates by email.
- Classifies records as Marketing Eligible, Transactional Only, Needs Review, or Do Not Contact.
- Exports clean CSV files and a markdown audit report.
- Uses a responsive layout optimized for desktop, smartphone-width windows, and PDA/small-screen views.
- Includes Electron Builder configuration for Windows `.exe` and macOS `.dmg` installers.
- Includes a GitHub Actions workflow to build installers on Windows and macOS runners.

## What this MVP intentionally does not do

- It does not scrape Etsy.
- It does not bypass account restrictions.
- It does not harvest credentials.
- It does not send bulk email.
- It does not mark contacts as marketing eligible unless the operator confirms or the file provides an opt-in-like field.

## Requirements

Recommended: Node.js 22 LTS or newer.

## Install and run locally

```bash
npm install
npm run dev
```

## Web-only preview

```bash
npm run web:dev
```

Then open the local Vite URL printed in your terminal.

## Build app assets

```bash
npm run build
npm start
```

## Package Windows EXE

```bash
npm run dist:win
```

## Package macOS DMG

```bash
npm run dist:mac
```

## Package current platform only

```bash
npm run dist
```

Installer output appears in:

```text
release/
```

## GitHub Actions installer build

The workflow file is included here:

```text
.github/workflows/build-installers.yml
```

Use GitHub Actions for the cleanest non-technical release flow:

- Windows runner produces `.exe`.
- macOS runner produces `.dmg`.

## Release build guide

See:

```text
RELEASE_BUILD_GUIDE.md
```

## Suggested next sprint

1. Add persistent local project storage with SQLite.
2. Add MBOX/EML parsing for inbox exports.
3. Add undo history for record edits.
4. Add PDF export for the audit report.
5. Add production icons: `.ico` for Windows and `.icns` for macOS.
6. Add signing/notarization pipeline before client-facing distribution.
