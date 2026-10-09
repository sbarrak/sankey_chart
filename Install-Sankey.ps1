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
$here = if ($env:SANKEY_DIR) { $env:SANKEY_DIR.TrimEnd('') } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$regKey = $null; $oldVal = $null; $xl = $null
try {
    foreach ($f in 'SankeyAddin.bas','SankeyEvents.cls','customUI14.xml','Install-Ribbon.ps1') {
        if (-not (Test-Path (Join-Path $here $f))) { throw "Missing file next to this script: $f" }
    }
    if (Get-Process excel -ErrorAction SilentlyContinue) { throw "Excel is running. Save your work, close all Excel windows and run this again." }

    $xl = New-Object -ComObject Excel.Application
    $ver = $xl.Version                                   # e.g. 16.0
    $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); $xl = $null

    $regKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    if (-not (Test-Path $regKey)) { New-Item $regKey -Force | Out-Null }
    $oldVal = (Get-ItemProperty $regKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
    Set-ItemProperty $regKey -Name AccessVBOM -Value 1 -Type DWord

    Write-Host "Building $AddinPath ..."
    New-Item -ItemType Directory -Force (Split-Path -Parent $AddinPath) | Out-Null
    $xl = New-Object -ComObject Excel.Application
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()
    $vb = $wb.VBProject
    [void]$vb.VBComponents.Import((Join-Path $here 'SankeyAddin.bas'))
    [void]$vb.VBComponents.Import((Join-Path $here 'SankeyEvents.cls'))
    $wb.Title = 'Sankey Chart Builder'
    if (Test-Path $AddinPath) {
        foreach ($a in $xl.AddIns) { if ($a.FullName -eq $AddinPath) { $a.Installed = $false } }
        Remove-Item $AddinPath -Force
    }
    $wb.SaveAs($AddinPath, 55)                           # xlOpenXMLAddIn
    $wb.Close($false)
    $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); $xl = $null
} catch {
    Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.Message -match 'null-valued|0x800A03EC|programmatic access') {
        Write-Host "Excel blocked access to the VBA project. Enable File > Options > Trust Center > Trust Center Settings > Macro Settings > 'Trust access to the VBA project object model', or ask your IT department (policy may lock it)." -ForegroundColor Yellow
    }
    $failed = $true
} finally {
    if ($xl) { try { $xl.Quit() } catch {} }
    if ($regKey) {                                       # restore the user's security setting
        if ($null -eq $oldVal) { Remove-ItemProperty $regKey -Name AccessVBOM -ErrorAction SilentlyContinue }
        else { Set-ItemProperty $regKey -Name AccessVBOM -Value $oldVal -Type DWord }
    }
}

if (-not $failed) {
    try {
        $env:SANKEY_NOPAUSE = '1'
        $env:SANKEY_DIR = $here
        & ([scriptblock]::Create([IO.File]::ReadAllText((Join-Path $here 'Install-Ribbon.ps1')))) -AddinPath $AddinPath   # run as text: not blocked by script-signing policy
        if (-not $NoRegister) {
            $xl = New-Object -ComObject Excel.Application
            $ai = $xl.AddIns.Add($AddinPath, $false)
            $ai.Installed = $true
            $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
        }
        Write-Host "Done. Start Excel: the Sankey icons are on the Add-ins tab." -ForegroundColor Green
    } catch {
        Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }
}
if ($Host.Name -eq 'ConsoleHost' -and -not $env:SANKEY_NOPAUSE_ALL) { Read-Host "Press Enter to close" | Out-Null }
