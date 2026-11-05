# Export every slide of a PPTX to PNG for page-by-page review.
# Usage:
#   powershell -NoProfile -File scripts/render_previews.ps1 -Pptx path/to/deck.pptx -OutDir render
# Optional: -Width 1600 (default 1600)
param(
  [Parameter(Mandatory = $true)][string]$Pptx,
  [Parameter(Mandatory = $false)][string]$OutDir = "render",
  [int]$Width = 1600
)
$ErrorActionPreference = "Stop"

$Pptx = (Resolve-Path -LiteralPath $Pptx).Path
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$OutDir = (Resolve-Path -LiteralPath $OutDir).Path

$ppt = New-Object -ComObject PowerPoint.Application
try {
  $pres = $ppt.Presentations.Open($Pptx, $true, $false, $false)
  $h = [int]([math]::Round($Width * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth))
  $n = $pres.Slides.Count
  for ($i = 1; $i -le $n; $i++) {
    $f = Join-Path $OutDir ("slide-{0:d2}.png" -f $i)
    $pres.Slides.Item($i).Export($f, "PNG", $Width, $h)
  }
  $pres.Close()
  Write-Output ("exported {0} slides ({1}x{2}) -> {3}" -f $n, $Width, $h, $OutDir)
}
finally {
  $ppt.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
