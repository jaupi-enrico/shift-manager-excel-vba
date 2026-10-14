# Status: branch `merge/no-ui-alerts`

Last updated: 2026-10-14. All planned code work is done and merged into this branch. Nothing has been run in Excel yet. No PR exists and none must be opened until the owner says so.

## Done (all merged into this branch)
- `fix/review-findings`: the code-review fixes (details below).
- `worktree-block-absent-workers`: `Lineare_Assenza` in `ModPulsanti`; `LineareLUN..DOM` reject a worker who is on FERIE/MALATTIA/CORSO that day. Now also on `MANAGER` (per 21-row block = weekday, via `Lineare_Assenza_Giorno` on `TOT-M`). Also the freeze-panes fix (skipped when the window is too narrow) in `GeneralWorkbook`.
- Block delete/paste (`Handle_Multi_Change`) on all seven day sheets and `MANAGER` (was only `LineareMAR`). Pasted names are also checked for absence.
- `obfuscate-password`: sheet password and login PIN are XOR-encoded and exposed as functions (`ModPulsanti`, `LoginForm`). They are no longer plain text in the repo.

## Why this branch exists
Every deployed workbook runs `Updater.AggiornaCodiceVBA` on open and pulls the modules from `main`, so anything merged to `main` ships to all users at once. All fixes from the code review live on this branch and are to be tested on a **copy** of `gestionale.xlsm` first.

## Branch state
- `main` is untouched. `code/1` and `code/2` are identical (checked with `diff -r`).
- Commits on the branch:
  - `3569f0a` error handling, roles, locale, bounded loops, helpers and constants in `ModPulsanti`
  - `aeeb35e` use the `SHEET_PASSWORD` / `LOGIN_PIN` constants
  - `864abb1` `Lineare*` modules: reset `OriginalRow`, multi-cell guard, message texts
  - `03d1993` updater hardening, `Workbook_Open` defers the update with `OnTime`, `Option Explicit` in `StampaFogli`
  - `1b60f04` `CLAUDE.md`

## IMPORTANT before merging to main
- `Updater.vba` (code/1 and code/2) has `UPDATE_BRANCH = "main"`, ready to merge.

## What was changed
| Area | Change |
|------|--------|
| `DatabaseWorkers*.vba` | `On Error GoTo Cleanup` restores events/screen updating. Multi-cell paste into the Ferie/Malattie/Corsi columns no longer fails (the Si/No fill only runs for a single cell). No-workers case exits cleanly. |
| `ModTotalone*.vba` | `Update_Workers` / `Update_Workers_M` have an error handler that restores app state and re-protects sheets. Unknown role (e.g. "N/A") falls back to Front (no more black rows). `Role_Color` helper. Bounded `Find_End_Row` replaces the `IMPRESA1` scans. `NameFound` is bounded. `Formula2` goes through `Set_Formula_Compat`. Conditional formats go through `Add_Cond_Format` (Italian syntax first, English fallback). |
| `Totalone*.vba` | `FormulaLocal "=SOMMA"` changed to `.Formula "=SUM"`. `Worksheet_Activate` re-protects the sheet on error. |
| `ModPulsanti.vba` | `SHEET_PASSWORD`, `LOGIN_PIN`, `RestoreAppState`, `Set_Formula_Compat`, `Formula_To_English`, `Add_Cond_Format` added here (new modules cannot be installed by the updater). `Delete_names` has an error handler, and the MANAGER branch skips rows without I and F. `Hide_Lines`: `clear_down_down` now checks `Row + 2`. `Hide_Lines` / `Show_Lines` turn off screen updating and restore state on error. |
| `Lineare*.vba` (8 files) | `OriginalRow` reset at the start and end of each event. A multi-cell edit shows "Modifica una cella alla volta" and exits. Message text fixes (già, è, "visualizzare"). |
| `Updater.vba` | Downloads all modules first and changes nothing if any download fails. Validates the downloaded text. Only replaces modules whose code differs. Exports a backup to `%TEMP%\shift-manager-backup-<timestamp>`. `Updater` is rewritten last. Branch is a constant. |
| `GeneralWorkbook.vba` | The update is scheduled with `Application.OnTime` (2 s) instead of being called inside `Workbook_Open`. |
| `StampaFogli.vba` | `Option Explicit` added. |
| `LoginForm`, `Main`, `Formazione*` and others | Use `SHEET_PASSWORD` / `LOGIN_PIN`. |

## Corrections to the review
- `BeforeClose` was correct (the `vbNo` branch exists). No change made.
- A blank contract cell converts to 0 and does not crash. Only non-numeric text raises the error, and that is now handled.
- The `.vba` files are UTF-8. The updater decodes UTF-8, so accents work. `CLAUDE.md` was corrected.

## What was NOT done
- Not run in Excel. The merged branches are also untested. Only block balance (`Sub`/`If`/`For`/`With`) was compared against the original files.
- The password and PIN are now obfuscated, not secure: anyone can decode them from the repo. Rotate them if that matters, since the old values stay in the commit history.
- `Lineare*` consolidation (8 near-identical modules into one) is not done. It is the riskiest change, so do it only after this round is tested.
- `Changes.ApplyChanges` is still never called.
- `ModTotaloneManager.vba`: the diff against `ModTotalone.vba` was only partly reviewed.
- The block handlers (`Handle_Multi_Change`) are the riskiest new code: test block delete and paste on a day sheet and on `MANAGER`. The `MANAGER` version uses the same logic as `LineareMAR` with its own ranges, and the duplicate-name check is limited to the weekday block, like the single-cell code.
- A module may be "different" every time if the VBE rewrites identifier casing, so it would be re-replaced on every open. This is no worse than before, but check it.

## Test plan (Windows, on a copy)
1. Copy `gestionale.xlsm` to `gestionale-test.xlsm` outside the repo. Do not open the deployed file.
2. In the VBE (Alt+F11), replace the `Updater` module with `code/2/Updater.vba` from this branch by hand (copy and paste). Enable "Trust access to the VBA project object model".
3. Open the copy with macros enabled. Expect "Codice aggiornato con successo! (N moduli)" and a backup folder in `%TEMP%`. Open it again: nothing should update, or only a few modules (see the casing note above).
4. Dipendenti: put text in the contract column, then run the update. Expect a clean error message with events/calculation/protection restored.
5. Dipendenti: use a role "N/A". Expect no black row.
6. Dipendenti: paste a block of cells into Ferie/Malattie/Corsi. Events must stay on (type in a day sheet afterwards to check).
7. Day sheets and MANAGER: edit I/F times, the 6:00 and 6:30 prompts, a duplicate name, Delete over a block (expect the message), `Hide_Lines` / `Show_Lines`.
8. MANAGER: `Delete_names` with a row that has a name but no I/F.
9. Close the workbook: Yes saves, No closes without saving, Cancel stays open.
10. Day sheets: enter a worker who is on FERIE/MALATTIA/CORSO that day (expect an error and the cell restored). On `LineareMAR`, block delete and paste.
11. If possible, run on a non-Italian Excel (locale fixes) and on Excel before 365 (`Formula2` fallback).

## Next steps
- Test on a copy and report results; fix whatever fails (on this branch).
- Decide on password/PIN, `Lineare*` consolidation, and `ApplyChanges`.
- When the owner says so: re-run `diff -r code/1 code/2` then open the PR.

## Code-review fixes (latest round)
- `Formula_To_English` is now a small tokenizer (E/O/SE/NON, `;`, decimal commas, quoted text untouched); callers skip the format if `Add_Cond_Format` returns `Nothing`.
- `DatabaseWorkers*`: separate `ErrHandler`, so a stale `Err` can no longer show a false error.
- `Updater`: phase 4 for the self-replacement, with an accurate message if only that step fails. `Workbook_Open` reports a failed `OnTime` scheduling.
- Checked and left as is: `Update_Validation`, `Add_Worker`, `Delete_Worker` are only called from `Update_Workers`, which has an error handler.
