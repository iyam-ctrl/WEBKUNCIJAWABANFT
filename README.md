# WEBKUNCIJAWABANFT

A responsive question-and-answer study app with a ChatGPT-like interface.

## Features
- 4 Days × 4 Parts = 16 modules.
- Chat-style question entry with answer choices in one input.
- Exact-match priority (100%) followed by similarity ranking.
- Search within the selected module, with optional expansion across modules.
- Shows the original source module and flags conflicting answers.
- Light/dark mode, responsive layout, and installable PWA shell.
- Excel source import (.xlsx/.xls) and IndexedDB persistence in the current browser.

## First-time setup
1. Open the app.
2. Choose **Impor materi Excel**.
3. Select the nine supplied source workbooks together.
4. The app maps files by filename to the corresponding topic modules.
5. Imported material is stored in this browser using IndexedDB.

## Important data note
The source workbooks are imported in the browser; this repository does not contain the original workbook contents. Browser IndexedDB is device/browser-local, so importing on one device does not automatically sync to another device. Use **Ekspor cadangan** and **Pulihkan cadangan** to move a data backup between devices.

## Deploy
Deploy this repository as a static site (for example, Vercel) with the repository root as the output directory. No build step is required.
