<#
    Export every slide of a PPTX to a 300 DPI PNG (A5 = 1754 x 2480 px).

    Usage:
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File export_png.ps1 `
            -Pptx "C:\work\sep2026_repaired.pptx" -OutDir "C:\work\png" -Log "C:\work\export_log.txt"

    IMPORTANT
    - Use ASCII-only paths. Chinese characters inside this script make PowerShell 5.1
      fail with HRESULT 0x8007007B. Copy the PPTX to a short ASCII name first.
    - COM output is often swallowed by the calling host, so everything is written to
      -Log; read that file afterwards instead of relying on stdout.
#>
param(
    [Parameter(Mandatory = $true)][string]$Pptx,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [string]$Log = "$PSScriptRoot\export_log.txt",
    [int]$Width = 1754,
    [int]$Height = 2480
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
Remove-Item "$OutDir\*.png" -ErrorAction SilentlyContinue

"opening $Pptx" | Out-File $Log -Encoding utf8
$ppt = New-Object -ComObject PowerPoint.Application
try {
    # ReadOnly=$true, Untitled=$false, WithWindow=$true
    $pres = $ppt.Presentations.Open($Pptx, $true, $false, $true)
    "SlideCount=$($pres.Slides.Count)" | Out-File $Log -Append -Encoding utf8

    foreach ($slide in $pres.Slides) {
        $n = $slide.SlideNumber.ToString('00')
        $slide.Export("$OutDir\slide_$n.png", "PNG", $Width, $Height)
        "exported slide_$n.png" | Out-File $Log -Append -Encoding utf8
    }

    $pres.Close()
} catch {
    "ERROR: $($_.Exception.Message)" | Out-File $Log -Append -Encoding utf8
    throw
} finally {
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}

"Done. Files: $((Get-ChildItem $OutDir -Filter *.png).Count)" | Out-File $Log -Append -Encoding utf8
