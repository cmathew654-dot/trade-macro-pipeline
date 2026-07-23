# Remove Live Market Data Integration

## Goal

Remove every current-code path that retrieves, refreshes, displays, tests, deploys, or documents external live market data in the CDS Trade Assistant.

The finished pipeline must not use Excel Stocks, linked data types, `ConvertToLinkedDataType`, `FIELDVALUE`, live quotes, price-refresh actions, or price-drift alerts.

## Critical Boundary

The local ticker-classification system remains intact.

The pipeline will continue to:

- Import ticker symbols and holdings from the advisor's custodial CSV.
- Read the static ticker-to-asset-class map stored on the local `CDS_Settings` worksheet.
- Classify holdings through `ClassifyTickerWithFallback`.
- Present unknown tickers for advisor review.
- Save advisor-approved classifications back to the local `CDS_Settings` worksheet.
- Reuse local classifications for holdings and buy-plan rows.

That process is self-referential and workbook-local. It does not retrieve external data and is explicitly outside the removal scope.

## Removal Scope

### Delete dedicated live-data artifacts

- Delete `vba/CDS_PriceGuard.bas`.
- Delete `tests/verify_price_guard.py`.
- Delete `deploy/Verify-LivePrices.ps1`.

### Remove runtime and user-interface wiring

- Remove the `refresh_prices` action from `vba/CDS_MacroLauncher.bas`.
- Remove the live-price callback from `vba/CDS_RibbonCallbacks.bas`.
- Remove the Refresh Live Prices control from `ribbon/customUI14.xml`.
- Remove the Refresh Live Prices button and state update from `vba/frmCDSTradeAssistant.frm`.
- Remove the `DriftAlertPct` setting from `vba/CDS_Settings.bas`.

### Remove residual references

- Remove the live-price stage, module description, entrypoint, safety explanation, test command, and related language from `README.md`.
- Remove live-price deployment steps from `deploy/README.md`.
- Rewrite snapshot comments in `vba/CDS_Snapshots.bas` so they describe freezing formulas and external links generically without referencing `FIELDVALUE` or linked data types.
- Update `deploy/Update-PersonalXlsb.ps1` so the next normal deployment removes the obsolete market-data module from an existing `PERSONAL.XLSB` before importing the remaining CDS modules.

## Preserved Functionality

The removal must not change:

- CSV import and holdings normalization.
- Local ticker and asset-class classification.
- Unknown-ticker review and persistence.
- Multi-account selection.
- Raise-cash scenarios.
- Sell instructions expressed as dollars, shares, percentage of position, percentage of account, or sell-all.
- Proceeds routing.
- Buy plans and local ticker classification inside buy plans.
- Tax, income, allocation, shortfall, routing, and wash-sale audits.
- Draft trade-email generation and routing refusal behavior.
- Values-only snapshots and snapshot export.
- Per-client session rollover.
- Macro launcher behavior unrelated to live pricing.
- Ribbon and assistant-form controls unrelated to live pricing.
- `PERSONAL.XLSB` deployment for the remaining CDS modules.

## Post-Removal Architecture

The workbook remains a fully local Excel/VBA pipeline.

Data flow:

1. A custodial CSV supplies static holdings values and ticker symbols.
2. The local `CDS_Settings` worksheet supplies ticker classifications and workflow settings.
3. Advisor review resolves unknown tickers locally.
4. Scenarios, sells, routing, buys, audits, emails, and snapshots operate only on imported or advisor-entered workbook data.

No component will request, convert, query, or refresh market information from an external source.

## Compatibility and Git Safety

- Existing Git history, tags, branch identities, and remote commits will not be rewritten.
- The removal will be implemented as new commits on the current branch.
- No force-push, rebase, filter-repo operation, or destructive remote action is permitted.
- The normal backed-up `PERSONAL.XLSB` deployment flow will remove the obsolete market-data module so an installed historical copy cannot remain callable.
- Existing workbooks may still contain an old `CDS Live Prices` worksheet created by earlier versions. The updated pipeline will ignore it and will not refresh, read, or recreate it.
- Existing `CDS_Settings` worksheets may retain a historical `DriftAlertPct` row. New source code will not create or read that setting. No migration macro is required.

## Test Strategy

### Negative boundary test

Add a source-level regression test that scans the current tracked runtime, Ribbon, deployment, test, and documentation files and fails if any banned live-data tokens remain:

- `CDS_PriceGuard`
- `RefreshLivePrices`
- `refresh_prices`
- `ConvertToLinkedDataType`
- `FIELDVALUE`
- `CDS Live Prices`
- `DriftAlertPct`
- `Excel Stocks`
- `linked data type`
- `live price`
- `live quote`

Historical Git objects are intentionally excluded because history is not being rewritten.

### Behavioral regression

Run the real headless Excel pipeline and all remaining focused regression scripts. The expected result is:

- VBA import and compilation succeed without `CDS_PriceGuard`.
- The full pipeline completes with zero audit failures.
- Local classification and unknown-ticker persistence continue to pass.
- All non-price workflow controls continue to resolve through the launcher, Ribbon, and assistant form.
- No repository source, current documentation, test, or deployment artifact contains a banned live-data reference.

## Acceptance Criteria

The change is complete when:

1. All dedicated live-data artifacts and wiring are removed.
2. The backed-up deployment updater prunes an installed historical market-data module.
3. The local static ticker-classification workflow is unchanged and tested.
4. The banned-token regression test passes.
5. The full Excel pipeline and every remaining focused regression script pass.
6. The working tree contains no untracked removal artifacts.
7. Git history and the remote remain untouched.
