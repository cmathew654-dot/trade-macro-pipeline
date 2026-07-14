# Deploying to PERSONAL.XLSB (work machine)

One-time update that syncs your production PERSONAL.XLSB to this repo's code.
Carry the whole kit folder (contains `vba\`, `ribbon\`, and these scripts) to
the work machine — the zip in `Backups\` already has everything.

## Steps

1. Close **all** Excel windows.
2. In PowerShell:

   ```powershell
   cd <kit folder>\deploy   # or the kit root if the scripts sit next to vba\
   powershell -ExecutionPolicy Bypass -File .\Update-PersonalXlsb.ps1
   ```

   What it does: backs up PERSONAL.XLSB next to itself
   (`PERSONAL-backup-<timestamp>.xlsb`), replaces every CDS module + the
   assistant form, replaces the embedded ribbon XML **only if one is already
   embedded**, and restores your Trust Center setting afterwards. Your
   `ThisWorkbook` code is never touched.

3. Open Excel, run the CDS launcher once to confirm it loads.
4. Verify live prices (the one thing the dev machine couldn't test):

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Verify-LivePrices.ps1
   ```

   `PASS` → RefreshLivePrices is fully live. `FAIL` with a license message →
   sign in to your M365 account inside Excel and re-run; until then the price
   guard uses its (tested) graceful-degradation path.

## Rollback

Copy the timestamped backup over PERSONAL.XLSB while Excel is closed.

## Notes

- If PERSONAL.XLSB lives somewhere non-standard: `-PersonalPath "C:\...\PERSONAL.XLSB"`.
- If your ribbon comes from a separate add-in (not embedded in PERSONAL.XLSB),
  the script says so — update that add-in's `customUI14.xml` from `ribbon\` instead.
- The scripts only ever terminate the hidden Excel instance they started.
