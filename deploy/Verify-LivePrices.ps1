<#
.SYNOPSIS
  Verifies that Excel Stocks linked data types actually fetch live prices
  on this machine - the one thing the dev machine could not test.

.DESCRIPTION
  Starts a hidden throwaway Excel, converts a few tickers with
  Range.ConvertToLinkedDataType (the exact call CDS_PriceGuard makes), asks
  for FIELDVALUE("Price"), and waits for real numbers.

  PASS  -> RefreshLivePrices in the CDS ribbon will work here.
  FAIL  -> the price guard will use its graceful-degradation path (already
           tested); fix the sign-in/license issue and re-run.

  Exit codes: 0 = PASS, 1 = timeout (converted but no price), 2 = data
  types unavailable on this Excel (license/sign-in).

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\Verify-LivePrices.ps1
  powershell -ExecutionPolicy Bypass -File .\Verify-LivePrices.ps1 -Visible
#>
param(
    [switch]$Visible,
    [int]$TimeoutSec = 90,
    [string[]]$Tickers = @('MSFT', 'AAPL')
)

$ErrorActionPreference = 'Stop'

Add-Type @'
using System;
using System.Runtime.InteropServices;
public class CdsWin32v {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
}
'@

$excel = $null; $excelPid = 0; $exitCode = 1
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$Visible
    $excel.DisplayAlerts = $false
    [void][CdsWin32v]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)

    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    for ($i = 0; $i -lt $Tickers.Count; $i++) {
        $ws.Cells.Item($i + 1, 1).Value2 = $Tickers[$i]
        $ws.Cells.Item($i + 1, 2).Formula = "=FIELDVALUE(A$($i + 1),""Price"")"
    }

    Write-Host "Converting $($Tickers -join ', ') to Stocks data types..."
    $converted = $true
    try {
        # ServiceID 268435456 = Stocks; same call CDS_PriceGuard.RefreshLivePrices uses
        $ws.Range("A1:A$($Tickers.Count)").ConvertToLinkedDataType(268435456, 'en-US')
    } catch {
        Write-Host ''
        Write-Host 'FAIL - this Excel cannot create Stocks data types.'
        Write-Host "  Error: $($_.Exception.Message)"
        Write-Host '  Usual causes: not signed in to a licensed M365 account, or the'
        Write-Host '  license lacks linked data types. Sign in inside Excel, then re-run.'
        Write-Host '  (This is the same condition CDS_PriceGuard degrades gracefully on.)'
        $exitCode = 2
        $converted = $false
    }

    if ($converted) {
    Write-Host "Waiting up to $TimeoutSec s for live prices..."
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    $prices = @{}
    while ((Get-Date) -lt $deadline -and $prices.Count -lt $Tickers.Count) {
        Start-Sleep -Seconds 2
        for ($i = 0; $i -lt $Tickers.Count; $i++) {
            $t = $Tickers[$i]
            if ($prices.ContainsKey($t)) { continue }
            $v = $ws.Cells.Item($i + 1, 2).Value2
            if ($v -is [double] -and $v -gt 0) { $prices[$t] = $v }
        }
    }

    Write-Host ''
    if ($prices.Count -eq $Tickers.Count) {
        foreach ($t in $Tickers) { Write-Host ("  {0,-6} {1,10:N2}" -f $t, $prices[$t]) }
        Write-Host ''
        Write-Host 'PASS - live prices fetched. RefreshLivePrices will work on this machine.'
        $exitCode = 0
    } else {
        $missing = $Tickers | Where-Object { -not $prices.ContainsKey($_) }
        Write-Host "FAIL - converted, but no price arrived for: $($missing -join ', ')."
        Write-Host '  Check network/proxy, then re-run with -Visible to look for prompts.'
        $exitCode = 1
    }
    }
} finally {
    if ($excel) {
        foreach ($w in @($excel.Workbooks)) { $w.Close($false) }
        $excel.Quit()
    }
    if ($excelPid) {
        $deadline = (Get-Date).AddSeconds(15)
        while ((Get-Date) -lt $deadline -and (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) { Start-Sleep -Milliseconds 500 }
        $p = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
        if ($p) { Stop-Process -Id $excelPid -Force }   # only the instance this script started
    }
}
exit $exitCode
