# Installing into a local PERSONAL.XLSB

Optional local setup for a user-owned Excel profile. This updater copies the
repository's code into your local PERSONAL.XLSB.

## Steps

1. Close **all** Excel windows.
2. Keep a backup of PERSONAL.XLSB before updating it.
3. From the `deploy` directory, run the script according to your machine's
   script-execution policy; do not bypass that policy:

   ```powershell
   .\Update-PersonalXlsb.ps1
   ```

   What it does: backs up PERSONAL.XLSB next to itself
   (`PERSONAL-backup-<timestamp>.xlsb`), prunes retired CDS modules, replaces
   every remaining CDS module + the assistant form, replaces the embedded
   ribbon XML **only if one is already embedded**, and restores your Trust
   Center setting afterwards. Your `ThisWorkbook` code is never touched.

4. Open Excel, run the CDS launcher once to confirm it loads.

## Rollback

Copy the timestamped backup over PERSONAL.XLSB while Excel is closed.

## Notes

- If PERSONAL.XLSB lives somewhere non-standard: `-PersonalPath "C:\...\PERSONAL.XLSB"`.
- If your ribbon comes from a separate add-in (not embedded in PERSONAL.XLSB),
  the script says so — update that add-in's `customUI14.xml` from `ribbon\` instead.
- The scripts only ever terminate the hidden Excel instance they started.
