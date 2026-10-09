<#
  One-step installer for the Sankey add-in (Windows, desktop Excel 2010+).
  Builds Sankey.xlam from SankeyAddin.bas + SankeyEvents.cls, adds the ribbon (customUI14.xml)
  and registers the add-in in Excel. Run it again any time to update.

  Double-click Install-Sankey.cmd, or:  .\Install-Sankey.ps1   [-AddinPath "C:\path\Sankey.xlam"] [-NoRegister]

  To import the code, Excel needs "Trust access to the VBA project object model". This script switches it
  on only while it builds the add-in and then puts your previous setting back.
#>
param(
    [string]$AddinPath = (Join-Path $env:APPDATA 'Microsoft\AddIns\Sankey.xlam'),
    [switch]$NoRegister
)
$ErrorActionPreference = 'Stop'
$here = if ($env:SANKEY_DIR) { $env:SANKEY_DIR.TrimEnd([char]92) } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$regKey = $null; $oldVal = $null; $xl = $null
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
try {
    foreach ($f in 'SankeyAddin.bas','SankeyEvents.cls','customUI14.xml') {
        if (-not (Test-Path (Join-Path $here $f))) { throw "Missing file in $here : $f" }
    }
    $tools = if ($env:SANKEY_TOOLS) { $env:SANKEY_TOOLS.TrimEnd([char]92) } else { $here }
    $ribbonScript = Join-Path $tools 'Install-Ribbon.ps1'
    if (-not (Test-Path $ribbonScript)) { throw "Missing file: $ribbonScript" }
    if (Get-Process excel -ErrorAction SilentlyContinue) { throw "Excel is running. Save your work, close all Excel windows and run this again." }

    $xl = New-Excel
    $ver = $xl.Version                                   # e.g. 16.0
    Close-Excel $xl; $xl = $null

    $regKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    if (-not (Test-Path $regKey)) { New-Item $regKey -Force | Out-Null }
    $oldVal = (Get-ItemProperty $regKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
    Set-ItemProperty $regKey -Name AccessVBOM -Value 1 -Type DWord

    Write-Host "Building $AddinPath ..."
    New-Item -ItemType Directory -Force (Split-Path -Parent $AddinPath) | Out-Null
    $xl = New-Excel
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()
    $vb = $wb.VBProject
    foreach ($m in Get-ChildItem $here -File | Where-Object { $_.Extension -in '.bas','.cls' } | Sort-Object Name) {
        [void]$vb.VBComponents.Import($m.FullName)         # every code module in the folder (Windows: 2, macOS build: 3)
    }
    $wb.Title = 'Sankey Chart Builder'
    if (Test-Path $AddinPath) {
        foreach ($a in $xl.AddIns) { if ($a.FullName -eq $AddinPath) { $a.Installed = $false } }
        Remove-Item $AddinPath -Force
    }
    $wb.SaveAs($AddinPath, 55)                           # xlOpenXMLAddIn
    $wb.Close($false); $wb = $null; $vb = $null
    Close-Excel $xl; $xl = $null
} catch {
    Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.Message -match 'null-valued|0x800A03EC|programmatic access') {
        Write-Host "Excel blocked access to the VBA project. Enable File > Options > Trust Center > Trust Center Settings > Macro Settings > 'Trust access to the VBA project object model', or ask your IT department (policy may lock it)." -ForegroundColor Yellow
    }
    $failed = $true
} finally {
    if ($xl) { Close-Excel $xl; $xl = $null }
    if ($regKey) {                                       # restore the user's security setting
        if ($null -eq $oldVal) { Remove-ItemProperty $regKey -Name AccessVBOM -ErrorAction SilentlyContinue }
        else { Set-ItemProperty $regKey -Name AccessVBOM -Value $oldVal -Type DWord }
    }
}

if (-not $failed) {
    try {
        $env:SANKEY_NOPAUSE = '1'
        $env:SANKEY_DIR = $here
        & ([scriptblock]::Create([IO.File]::ReadAllText($ribbonScript))) -AddinPath $AddinPath -Force   # run as text: not blocked by script-signing policy
        if (-not $NoRegister) {
            $xl = New-Excel
            $xl.DisplayAlerts = $false
            [void]$xl.Workbooks.Add()                    # AddIns.Add needs an open workbook
            $ai = $xl.AddIns.Add($AddinPath, $false)
            $ai.Installed = $true
            Close-Excel $xl; $xl = $null
        }
        Write-Host "Done. Start Excel: the Sankey icons are on the Add-ins tab." -ForegroundColor Green
    } catch {
        Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }
}
if ($Host.Name -eq 'ConsoleHost' -and -not $env:SANKEY_NOPAUSE_ALL) { Read-Host "Press Enter to close" | Out-Null }
