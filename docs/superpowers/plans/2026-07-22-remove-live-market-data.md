# Remove Live Market Data Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Remove all current Excel Stocks and external live-market-data functionality while preserving the workbook-local ticker classifier and every other CDS workflow.

**Architecture:** Delete the dedicated price subsystem, remove every invocation and user-facing control, and add a negative source-boundary regression that prevents reintroduction. Teach the backed-up PERSONAL.XLSB updater to prune the obsolete installed module before importing the remaining local-only modules.

**Tech Stack:** Excel VBA, Ribbon XML, PowerShell, Python 3, pywin32, real headless Microsoft Excel.

## Global Constraints

- Do not rewrite Git history, tags, branches, or remote commits.
- Do not use force-push, rebase, filter-repo, or destructive remote operations.
- Preserve CDS_Settings, ClassifyTickerWithFallback, unknown-ticker review, and local classification persistence.
- Preserve CSV import, scenarios, sells, routing, buys, audits, email drafts, snapshots, session rollover, and remaining UI behavior.
- Do not execute the production PERSONAL.XLSB updater during development verification.
- Keep VBA text CRLF and the form resource binary.

---

### Task 1: Add the absence regression and prove RED

**Files:**
- Create: `C:\Users\Cyril\Projects\cds-trade-assistant\tests\verify_no_live_market_data.py`

**Interfaces:**
- Consumes: current files under vba, ribbon, deploy, tests, and the root README.
- Produces: exit code 0 only when forbidden live-data tokens are absent and deployment cleanup exists.

- [ ] **Step 1: Create the failing regression**

```python
"""Regression boundary: the current CDS pipeline contains no live-market-data integration."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCAN_ROOTS = ("vba", "ribbon", "deploy", "tests")
TEXT_SUFFIXES = {".bas", ".cls", ".frm", ".xml", ".ps1", ".py", ".md"}
BANNED = tuple("".join(parts).lower() for parts in (
    ("CDS_", "PriceGuard"), ("Refresh", "LivePrices"),
    ("refresh", "_prices"), ("ConvertTo", "LinkedDataType"),
    ("FIELD", "VALUE"), ("CDS Live", " Prices"),
    ("DriftAlert", "Pct"), ("Excel ", "Stocks"),
    ("linked data", " type"), ("live ", "price"), ("live ", "quote"),
))

def candidate_files():
    files = [ROOT / "README.md"]
    for directory in SCAN_ROOTS:
        files.extend(path for path in (ROOT / directory).rglob("*")
                     if path.is_file() and path.suffix.lower() in TEXT_SUFFIXES)
    return sorted(files)

def main():
    violations = []
    for path in candidate_files():
        text = path.read_text(encoding="utf-8", errors="replace").lower()
        for token in BANNED:
            if token in text:
                violations.append(f"{path.relative_to(ROOT)}: {token}")
    updater = (ROOT / "deploy" / "Update-PersonalXlsb.ps1").read_text(
        encoding="utf-8", errors="replace")
    for required in ("$obsoleteModuleNames", "$proj.VBComponents.Remove($obsolete)"):
        if required not in updater:
            violations.append(f"deploy/Update-PersonalXlsb.ps1: missing {required}")
    if violations:
        print("FAIL: live-market-data boundary violations")
        for violation in violations:
            print(f"  {violation}")
        return 1
    print("PASS: no live-market-data integration remains in the current pipeline")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 2: Verify RED**

Run: `python C:\Users\Cyril\Projects\cds-trade-assistant\tests\verify_no_live_market_data.py`

Expected: exit 1 with current module, launcher, UI, setting, verifier, docs, comments, and missing updater cleanup violations.

---

### Task 2: Remove runtime, UI, setting, and dedicated artifacts

**Files:**
- Delete: `C:\Users\Cyril\Projects\cds-trade-assistant\vba\CDS_PriceGuard.bas`
- Delete: `C:\Users\Cyril\Projects\cds-trade-assistant\tests\verify_price_guard.py`
- Delete: `C:\Users\Cyril\Projects\cds-trade-assistant\deploy\Verify-LivePrices.ps1`
- Modify: `C:\Users\Cyril\Projects\cds-trade-assistant\vba\CDS_MacroLauncher.bas`
- Modify: `C:\Users\Cyril\Projects\cds-trade-assistant\vba\CDS_RibbonCallbacks.bas`
- Modify: `C:\Users\Cyril\Projects\cds-trade-assistant\ribbon\customUI14.xml`
- Modify: `C:\Users\Cyril\Projects\cds-trade-assistant\vba\frmCDSTradeAssistant.frm`
- Modify: `C:\Users\Cyril\Projects\cds-trade-assistant\vba\CDS_Settings.bas`

**Interfaces:**
- Consumes: failing boundary regression.
- Produces: no callable or visible external market-data behavior.

- [ ] Delete the three dedicated artifacts.
- [ ] Delete all four refresh_prices branches/labels from CDS_MacroLauncher.bas.
- [ ] Delete RunRefreshPrices from CDS_RibbonCallbacks.bas.
- [ ] Delete btnRefreshPrices from customUI14.xml; preserve both snapshot buttons.
- [ ] Delete cmdRefreshPrices creation and state update from the form.
- [ ] Move Save Snapshot and Export Snapshot buttons from top 720 to top 686 to close the gap.
- [ ] Delete the DriftAlertPct WriteSetting row; do not alter ticker-map code.
- [ ] Run the absence regression. Expected: still RED only for deployment cleanup, docs, and snapshot comments.

---

### Task 3: Prune the obsolete installed module safely

**Files:**
- Modify: `C:\Users\Cyril\Projects\cds-trade-assistant\deploy\Update-PersonalXlsb.ps1`

**Interfaces:**
- Consumes: an existing backed-up PERSONAL.XLSB VBA project.
- Produces: removal of the obsolete module before remaining sources import.

- [ ] Change `$replaced = @(); $added = @()` to `$replaced = @(); $added = @(); $removed = @()`.
- [ ] After acquiring `$proj`, add:

```powershell
$obsoleteModuleNames = @('CDS_' + 'PriceGuard')
foreach ($obsoleteName in $obsoleteModuleNames) {
    $obsolete = $null
    foreach ($comp in @($proj.VBComponents)) {
        if ($comp.Name -eq $obsoleteName) { $obsolete = $comp; break }
    }
    if ($obsolete) {
        $proj.VBComponents.Remove($obsolete)
        $removed += $obsoleteName
    }
}
```

- [ ] Add `Write-Host ("Removed  ({0}): {1}" -f $removed.Count, ($removed -join ', '))` beside Replaced/Added output.
- [ ] Replace the final instruction with `Write-Host 'Open Excel and run the CDS launcher to confirm the remaining pipeline.'`.
- [ ] Do not execute this production updater during development.

---

### Task 4: Remove documentation and source-comment residue

**Files:**
- Modify: `C:\Users\Cyril\Projects\cds-trade-assistant\README.md`
- Modify: `C:\Users\Cyril\Projects\cds-trade-assistant\deploy\README.md`
- Modify: `C:\Users\Cyril\Projects\cds-trade-assistant\vba\CDS_Snapshots.bas`

**Interfaces:**
- Consumes: local-only runtime from Tasks 2-3.
- Produces: documentation describing only imported/advisor-entered data.

- [ ] Remove price checks from the opening feature list, Mermaid flow, module list, entrypoints, safety model, and test commands.
- [ ] State: `Runs locally inside Excel; no backend, external API calls, or market-data retrieval. All calculations use values imported from the custodial CSV or entered by the advisor.`
- [ ] Remove the old deployment verification step and document obsolete-module pruning.
- [ ] Replace snapshot comments with:

```vb
'   Formulas and workbook links make the working sheet mutable; a
'   snapshot is a deliberate, non-live copy.
```

```vb
' Freezes every formula and workbook-linked value to a plain value.
' Run twice so copied cells are fully detached from mutable workbook state.
```

- [ ] Run the absence regression. Expected: PASS.
- [ ] Commit with `git commit -m "refactor: remove external live market data"`.

---

### Task 5: Verify the preserved classifier and full pipeline

**Files:**
- Test: all Python scripts under `C:\Users\Cyril\Projects\cds-trade-assistant\tests`
- Verify: `C:\Users\Cyril\Projects\cds-trade-assistant\vba\CDS_Settings.bas`
- Verify: `C:\Users\Cyril\Projects\cds-trade-assistant\vba\CDS_Unknowns.bas`

**Interfaces:**
- Consumes: cleaned current tree.
- Produces: evidence for local classification and all retained workflows.

- [ ] Run `python C:\Users\Cyril\Projects\cds-trade-assistant\tests\verify_no_live_market_data.py`; expect PASS.
- [ ] Run `python C:\Users\Cyril\Projects\cds-trade-assistant\tests\run_pipeline.py`; expect every step and zero audit failures.
- [ ] Run verify_amount_spec.py, verify_routing.py, verify_guards.py, verify_wash_flag.py, verify_email.py, verify_snapshots.py, and verify_new_session.py; expect OVERALL: PASS from each.
- [ ] Run `rg -n "ClassifyTickerWithFallback|AddTickerClassToSettings|SaveUnknownsAndRefresh" C:\Users\Cyril\Projects\cds-trade-assistant\vba`; expect settings, unknown-review, holdings, and buy-plan matches.
- [ ] Run `git -C C:\Users\Cyril\Projects\cds-trade-assistant diff --check` and `git -C C:\Users\Cyril\Projects\cds-trade-assistant status --short --branch --untracked-files=all`; expect no whitespace errors and only intentional changes before commit or a clean tree after commit.
