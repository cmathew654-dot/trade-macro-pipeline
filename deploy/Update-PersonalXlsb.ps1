<#
.SYNOPSIS
  Updates PERSONAL.XLSB with the CDS Trade Assistant modules from this kit.

.DESCRIPTION
  Backs up PERSONAL.XLSB, then replaces every CDS module (.bas), the
  CDS_ButtonHandler class, and the assistant form with the versions shipped
  next to this script. If a ribbon (customUI14.xml) is embedded in
  PERSONAL.XLSB it is replaced too; if not, that step is skipped.

  ThisWorkbook.cls is never touched (document module - yours stays as is).

  CLOSE ALL EXCEL WINDOWS before running. The script refuses to run while
  PERSONAL.XLSB is locked by another process, and it only ever terminates
  the hidden Excel instance it started itself.

.EXAMPLE
  .\Update-PersonalXlsb.ps1
#>
param(
    [string]$PersonalPath = (Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\PERSONAL.XLSB'),
    [string]$SourceRoot = '',
    [string]$OfficeVersion = '16.0'
)

$ErrorActionPreference = 'Stop'

# --- locate kit sources (works from deploy\ inside the repo or from kit root)
if (-not $SourceRoot) {
    foreach ($candidate in @($PSScriptRoot, (Split-Path $PSScriptRoot -Parent))) {
        if (Test-Path (Join-Path $candidate 'vba\frmCDSTradeAssistant.frm')) { $SourceRoot = $candidate; break }
    }
}
$vbaDir    = Join-Path $SourceRoot 'vba'
$ribbonXml = Join-Path $SourceRoot 'ribbon\customUI14.xml'
if (-not (Test-Path (Join-Path $vbaDir 'frmCDSTradeAssistant.frm'))) {
    throw "Cannot find vba\ sources under '$SourceRoot'. Pass -SourceRoot <kit folder>."
}
if (-not (Test-Path $PersonalPath)) {
    throw "PERSONAL.XLSB not found at '$PersonalPath'. Pass -PersonalPath explicitly."
}

# --- refuse to run while the file is open in another Excel
try {
    $probe = [System.IO.File]::Open($PersonalPath, 'Open', 'ReadWrite', 'None')
    $probe.Close()
} catch {
    throw "PERSONAL.XLSB is locked - close ALL Excel windows and re-run."
}

# --- backup
$stamp  = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = Join-Path (Split-Path $PersonalPath -Parent) "PERSONAL-backup-$stamp.xlsb"
Copy-Item $PersonalPath $backup
Write-Host "Backup: $backup"

# --- ensure VBE object model access (restored afterwards)
$secKey  = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Excel\Security"
$oldVbom = $null
if (Test-Path $secKey) { $oldVbom = (Get-ItemProperty $secKey -ErrorAction SilentlyContinue).AccessVBOM }
if ($oldVbom -ne 1) {
    if (-not (Test-Path $secKey)) { New-Item $secKey -Force | Out-Null }
    Set-ItemProperty $secKey -Name AccessVBOM -Value 1 -Type DWord
    Write-Host 'Trust access to VBA object model: temporarily enabled.'
}

Add-Type @'
using System;
using System.Runtime.InteropServices;
public class CdsWin32 {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
}
'@

function Get-ModuleName([string]$file) {
    foreach ($line in (Get-Content $file -TotalCount 15)) {
        if ($line -match 'Attribute VB_Name = "(.+)"') { return $Matches[1] }
    }
    return [System.IO.Path]::GetFileNameWithoutExtension($file)
}

$excel = $null; $excelPid = 0
$replaced = @(); $added = @(); $removed = @()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    [void][CdsWin32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)

    $wb = $excel.Workbooks.Open($PersonalPath)
    try { $proj = $wb.VBProject } catch {
        throw 'Cannot access the VBA project. Enable File > Options > Trust Center > Macro Settings > "Trust access to the VBA project object model", then re-run.'
    }
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


    # everything except the document module
    $sources  = @(Get-ChildItem (Join-Path $vbaDir '*.bas'))
    $sources += Get-ChildItem (Join-Path $vbaDir '*.cls') | Where-Object { $_.Name -ne 'ThisWorkbook.cls' }
    $sources += Get-ChildItem (Join-Path $vbaDir '*.frm')

    foreach ($src in $sources) {
        $name = Get-ModuleName $src.FullName
        $existing = $null
        foreach ($comp in @($proj.VBComponents)) { if ($comp.Name -eq $name) { $existing = $comp } }
        if ($existing) { $proj.VBComponents.Remove($existing); $replaced += $name }
        else           { $added += $name }
        [void]$proj.VBComponents.Import($src.FullName)
    }

    $wb.Save()
    $wb.Close($false)
} finally {
    if ($excel) { $excel.Quit() }
    if ($excelPid) {
        $deadline = (Get-Date).AddSeconds(15)
        while ((Get-Date) -lt $deadline -and (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) { Start-Sleep -Milliseconds 500 }
        $p = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
        if ($p) { Stop-Process -Id $excelPid -Force }   # only the instance this script started
    }
    if ($oldVbom -ne 1) {
        if ($null -eq $oldVbom) { Remove-ItemProperty $secKey -Name AccessVBOM -ErrorAction SilentlyContinue }
        else { Set-ItemProperty $secKey -Name AccessVBOM -Value $oldVbom -Type DWord }
        Write-Host 'Trust access to VBA object model: restored to previous setting.'
    }
}

# --- ribbon: replace embedded customUI14.xml only if one already exists
$ribbonStatus = 'no ribbon\customUI14.xml in kit - skipped'
if (Test-Path $ribbonXml) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($PersonalPath, 'Update')
    try {
        $entry = $zip.Entries | Where-Object { $_.FullName -match 'customUI14\.xml$' } | Select-Object -First 1
        if ($entry) {
            $full = $entry.FullName
            $entry.Delete()
            $new    = $zip.CreateEntry($full)
            $bytes  = [System.IO.File]::ReadAllBytes($ribbonXml)
            $stream = $new.Open()
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Close()
            $ribbonStatus = "replaced embedded $full"
        } else {
            $ribbonStatus = 'PERSONAL.XLSB has no embedded customUI part - skipped (if your ribbon comes from a separate add-in, update it there)'
        }
    } finally { $zip.Dispose() }
}

Write-Host ''
Write-Host '=== UPDATE COMPLETE ==='
Write-Host ("Replaced ({0}): {1}" -f $replaced.Count, ($replaced -join ', '))
Write-Host ("Added    ({0}): {1}" -f $added.Count, ($added -join ', '))
Write-Host ("Removed  ({0}): {1}" -f $removed.Count, ($removed -join ', '))
Write-Host "Ribbon: $ribbonStatus"
Write-Host "Backup: $backup"
Write-Host 'Open Excel and run the CDS launcher to confirm the remaining pipeline.'
