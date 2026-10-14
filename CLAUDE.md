# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

An Excel workbook (`gestionale.xlsm`) that manages employee shifts for a McDonald's-linked restaurant, driven entirely by VBA. There is no build, lint, or test tooling: the repo holds the workbook plus its VBA modules exported as plain text (`code/<n>/*.vba`). Code and UI strings are in Italian (LUN/MAR/MER/GIO/VEN/SAB/DOM = Mon–Sun, `Dipendenti`/`DatabaseWorkers` = employees, `Formazione` = training, `Totalone` = weekly recap, `Lineare` = per-day linear shift sheet, `Stampa` = print).

## Workflow and gotchas

- The `.vba` files are **not** compiled from the repo. Edits are made in the Excel VBE (or in these files) and must be copied into the workbook's modules by hand; the `.xlsm` is the real artifact. Keep the `.vba` exports in sync with the workbook.
- **`code/1/` and `code/2/` are currently identical.** The number is a "version/type" selector: `Updater.vba` reads `Worksheets("Tabelle").Range("H9")` and downloads `https://raw.githubusercontent.com/jaupi-enrico/shift-manager-excel-vba/main/code/<H9>/<ModuleName>.vba` for every VBA component in the workbook. When changing behavior, apply it to whichever variant(s) are meant to ship, and keep module file names equal to the VBA component names, because the updater matches them by name.
- Self-update (`AggiornaCodiceVBA` in `Updater.vba`) overwrites each module's code via `VBProject.VBComponents` (needs "Trust access to the VBA project object model"; error 1004 is handled with a hint). It strips a UTF-8 BOM and skips modules that return 404. `Changes.vba` (`ApplyChanges`) is a hook for one-off post-update migrations.
- The `.vba` files are UTF-8; the updater decodes them as UTF-8 and the VBE converts to ANSI on insert, so accented characters in strings work. Do not write `?` in place of accents.
- The sheet-protection password and login PIN are the constants `SHEET_PASSWORD` / `LOGIN_PIN` at the top of `ModPulsanti.vba`; shared helpers (`RestoreAppState`, `Set_Formula_Compat`, `Add_Cond_Format`) live there too.
- The updater only updates modules that already exist in the deployed workbook; a new `.vba` file is never installed. Put new shared code into an existing module (e.g. `ModPulsanti`).
- `Updater.vba` has `UPDATE_BRANCH`. It must be `"main"` on `main`; use another branch only for test copies of the workbook and revert it before merging.
- Several standard modules define the same procedure names (e.g. `Add_Worker` in `ModTotalone` and `ModTotaloneManager`). Calling one unqualified from a third module is an "Ambiguous name" error, so keep helper names unique.

## Architecture

Module naming maps onto workbook sheets and shared logic:

- **Per-day sheet modules** `LineareLUN/MAR/MER/GIO/VEN/SAB/DOM` are the code-behind of the seven day sheets. They are ~610 lines each and near-duplicates (LUN is slightly longer), so a fix usually must be applied to all seven, and `LineareManager`, which is the same logic for the `MANAGER` sheet.
- **`Totalone` / `TotaloneManager` / `ModTotalone` / `ModTotaloneManager`**: the weekly recap sheet (`TOT`) and its manager variant (`TOT-M`). `Mod*` hold the logic (e.g. `Update_Workers_M`, name lookup in `TOT-M`); the others are sheet code-behind.
- **`DatabaseWorkers` / `DatabaseWorkersManager`**: the employee database sheets (contracts, roles, pauses); the `Manager` variants mirror the crew ones for the managers' sheets.
- **`Formazione` / `FormazioneManager`**: training-day tracking, with absences (days off, sick days, formation days) highlighted visually.
- **`GeneralWorkbook`**: the `ThisWorkbook` module. `Workbook_BeforeClose` asks to save, then protects every sheet and activates the day/`MANAGER` sheets.
- **`Main`**: a sheet code-behind that auto-protects the sheet on activation and unprotects only the input range (`C21:J26`, `B2`) when selected.
- **`ModPulsanti`**: handlers for the dashboard buttons. **`StampaFogli`**: printing. **`LoginForm`**, `Update` (UserForm shown during update): forms.

Shared pattern in the `Linear*` modules: module-level `Dim` globals hold the working state (`rngName`, `ColumnName`, `OriginalRow`, `columnI/columnF`, …) and helper functions like `Find_Name_Column` place a worker's name in a free column of the shift grid, searching forward then backward and treating empty/pause cells as unavailable. Each file uses `Option Explicit`.

Recent commits touch column-A-based last-row detection, `TOT-M` sheet references in name lookup, and `codeModule` casing in the updater, so these are the active areas.
