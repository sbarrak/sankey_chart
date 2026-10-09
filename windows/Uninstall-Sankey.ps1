<#
  Removes the Sankey add-in from Excel: unloads it, deletes its ribbon (it lives inside the add-in file),
  removes its entries from Excel's add-in list and deletes Sankey.xlam (and the .bak made by the installer).

  Double-click Uninstall-Sankey.cmd, or:  .\Uninstall-Sankey.ps1 [-AddinPath "C:\path\Sankey.xlam"] [-Yes]

  Not touched: charts you already drew (they stay in your workbooks as ordinary shapes) and your data.
  Auto-update stops because the add-in is gone; delete the charts yourself if you do not want them.
#>
param(
    [string]$AddinPath = (Join-Path $env:APPDATA 'Microsoft\AddIns\Sankey.xlam'),
    [switch]$Yes
)
$ErrorActionPreference = 'Stop'
$failed = $false
$paths = New-Object System.Collections.Generic.List[string]
$paths.Add($AddinPath)

Add-Type -Namespace Sankey -Name Win -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint pid);'
$script:xlPid = 0
function New-Excel {                              # start a hidden Excel and remember ITS process id (never touches other Excel windows)
    $app = New-Object -ComObject Excel.Application
    $id = [uint32]0
    try { [void][Sankey.Win]::GetWindowThreadProcessId([IntPtr]$app.Hwnd, [ref]$id) } catch {}
    $script:xlPid = [int]$id
    return $app
}
function Close-Excel($app) {                      # quit our hidden Excel for real, even if COM references linger
    try { if ($app) { $app.Quit() } } catch {}
    if ($app) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    if ($script:xlPid -gt 0) {
        for ($i = 0; $i -lt 20 -and (Get-Process -Id $script:xlPid -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Milliseconds 500 }
        Get-Process -Id $script:xlPid -ErrorAction SilentlyContinue | Stop-Process -Force   # only the instance this script started
    }
}

$xl = $null
try {
    if (Get-Process excel -ErrorAction SilentlyContinue) { throw "Excel is running. Save your work, close all Excel windows and run this again." }

    if (-not $Yes) {
        $a = Read-Host "Remove the Sankey add-in and its ribbon from Excel? (y/N)"
        if ($a -notmatch '^(y|yes)$') { throw "Cancelled - nothing was changed." }
    }

    # 1. Unload through Excel itself (also finds copies installed in another folder)
    try {
        $xl = New-Excel
        $xl.DisplayAlerts = $false
        [void]$xl.Workbooks.Add()                  # the AddIns collection needs an open workbook
        foreach ($ai in @($xl.AddIns)) {
            if ($ai.Name -ieq 'Sankey.xlam' -or $ai.FullName -ieq $AddinPath) {
                if ($ai.FullName -and -not $paths.Contains($ai.FullName)) { $paths.Add($ai.FullName) }
                if ($ai.Installed) { $ai.Installed = $false; Write-Host "Unloaded: $($ai.FullName)" }
            }
        }
    } catch {
        Write-Host "Note: could not unload through Excel ($($_.Exception.Message)); cleaning up directly." -ForegroundColor Yellow
    } finally {
        if ($xl) { Close-Excel $xl; $xl = $null }
    }

    # 2. Remove Excel's remembered entries: auto-open list (OPEN, OPEN1, ...) and the Add-in Manager list
    foreach ($ver in Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match '^\d+\.\d+$' }) {
        $opt = "$($ver.PSPath)\Excel\Options"
        if (Test-Path $opt) {
            $props = Get-ItemProperty $opt
            foreach ($n in $props.PSObject.Properties.Name) {
                if ($n -match '^OPEN\d*$' -and "$($props.$n)" -match 'Sankey\.xlam') {
                    Remove-ItemProperty $opt -Name $n -ErrorAction SilentlyContinue; Write-Host "Removed registry entry $n"
                }
            }
        }
        $mgr = "$($ver.PSPath)\Excel\Add-in Manager"
        if (Test-Path $mgr) {
            $props = Get-ItemProperty $mgr
            foreach ($n in $props.PSObject.Properties.Name) {
                if ($n -match 'Sankey\.xlam$') { Remove-ItemProperty $mgr -Name $n -ErrorAction SilentlyContinue; Write-Host "Removed add-in list entry" }
            }
        }
    }

    # 3. Delete the add-in file(s) and the installer's backup
    $removed = 0
    foreach ($p in $paths) {
        foreach ($f in $p, "$p.bak") {
            if (Test-Path $f) {
                for ($i = 0; $i -lt 5; $i++) {
                    try { Remove-Item $f -Force; $removed++; Write-Host "Deleted $f"; break }
                    catch { if ($i -eq 4) { throw "Cannot delete $f - is Excel (or another program) still using it?" }; Start-Sleep -Seconds 1 }
                }
            }
        }
    }
    if ($removed -eq 0) { Write-Host "No Sankey add-in file found at $AddinPath (nothing to delete)." }

    Write-Host "Done. Sankey is removed. Restart Excel to confirm the Add-ins tab no longer shows the Sankey buttons." -ForegroundColor Green
} catch {
    Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
    $failed = $true
} finally {
    if ($xl) { Close-Excel $xl }
}
if ($Host.Name -eq 'ConsoleHost' -and -not $env:SANKEY_NOPAUSE) { Read-Host "Press Enter to close" | Out-Null }
