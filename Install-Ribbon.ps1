<#
  Adds the "Sankey" ribbon tab (customUI14.xml) to the saved add-in file.
  1. Save the add-in as Sankey.xlam first (see README).
  2. CLOSE Excel.
  3. Right-click this file > Run with PowerShell   (or:  .\Install-Ribbon.ps1 -AddinPath "C:\path\Sankey.xlam")
  A backup (Sankey.xlam.bak) is created next to the add-in.
#>
param(
    [string]$AddinPath = (Join-Path $env:APPDATA 'Microsoft\AddIns\Sankey.xlam'),
    [switch]$Force
)
$ErrorActionPreference = 'Stop'
try {
    $xmlFile = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'customUI14.xml'
    if (-not (Test-Path $xmlFile))   { throw "customUI14.xml not found next to this script." }
    if (-not (Test-Path $AddinPath)) { throw "Add-in not found: $AddinPath`nSave it as Sankey.xlam first, or pass -AddinPath." }
    if (-not $Force -and (Get-Process excel -ErrorAction SilentlyContinue)) { throw "Excel is running. Close all Excel windows and run this again." }

    $xml = [IO.File]::ReadAllText($xmlFile, [Text.Encoding]::UTF8)
    [void][xml]$xml                                    # fail early if the XML is malformed

    Copy-Item $AddinPath "$AddinPath.bak" -Force
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $utf8 = New-Object Text.UTF8Encoding($false)

    $zip = [IO.Compression.ZipFile]::Open($AddinPath, 'Update')
    try {
        $old = $zip.GetEntry('customUI/customUI14.xml'); if ($old) { $old.Delete() }
        $entry = $zip.CreateEntry('customUI/customUI14.xml')
        $w = New-Object IO.StreamWriter($entry.Open(), $utf8); $w.Write($xml); $w.Close()

        $relEntry = $zip.GetEntry('_rels/.rels')
        if (-not $relEntry) { throw "Not a valid Office file (_rels/.rels missing)." }
        $r = New-Object IO.StreamReader($relEntry.Open(), [Text.Encoding]::UTF8); $rels = $r.ReadToEnd(); $r.Close()
        if ($rels -notmatch 'customUI/customUI14\.xml') {
            $rel = '<Relationship Id="rIdSankeyUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
            $rels = $rels -replace '</Relationships>', ($rel + '</Relationships>')
            $relEntry.Delete()
            $ne = $zip.CreateEntry('_rels/.rels')
            $w = New-Object IO.StreamWriter($ne.Open(), $utf8); $w.Write($rels); $w.Close()
        }
    } finally { $zip.Dispose() }

    Write-Host "Ribbon installed into $AddinPath" -ForegroundColor Green
    Write-Host "Start Excel - you should see a 'Sankey' tab."
} catch {
    Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
}
if ($Host.Name -eq 'ConsoleHost' -and -not $env:SANKEY_NOPAUSE) { Read-Host "Press Enter to close" | Out-Null }
