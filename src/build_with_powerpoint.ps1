
param(
    [string]$Template  = "D:\work\ai-ppt-workflow\templates\xauat.pptx",
    [string]$Content   = "D:\work\ai-ppt-workflow\input\xauat-content-text.json",
    [string]$Out       = "D:\work\ai-ppt-workflow\output\nylon66-xauat.pptx",
    [string]$RenderDir = "D:\work\ai-ppt-workflow\render"
)
$ErrorActionPreference = 'Stop'
$cfg = Get-Content -LiteralPath $Content -Raw -Encoding UTF8 | ConvertFrom-Json

# Walk every shape (recursing into groups) and hand it to the callback.
function Walk($shape, $map) {
    if ($shape.Type -eq 6) { foreach ($c in $shape.GroupItems) { Walk $c $map }; return }
    if (-not $shape.HasTextFrame) { return }
    if (-not $shape.TextFrame.HasText) { return }
    $tr = $shape.TextFrame.TextRange
    $txt = $tr.Text
    if ([string]::IsNullOrEmpty($txt)) { return }

    # exact match first, then "starts with" so multi-paragraph shapes match
    $hit = $null
    foreach ($old in $map.PSObject.Properties.Name) {
        if ($txt -eq $old) { $hit = $old; break }
    }
    if ($null -eq $hit) {
        foreach ($old in $map.PSObject.Properties.Name) {
            if ($txt.StartsWith($old)) { $hit = $old; break }
        }
    }
    if ($null -ne $hit) {
        $new = [string]$map.$hit
        if ($new -ne $txt) { $tr.Text = $new }
    }
}

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($Template, $false, $false, $false)
try {
    foreach ($entry in $cfg.slides) {
        $sld = $pres.Slides.Item([int]$entry.src)
        foreach ($shape in $sld.Shapes) { Walk $shape $entry.map }
    }
    $keep = @($cfg.slides | ForEach-Object { [int]$_.src })
    for ($i = $pres.Slides.Count; $i -ge 1; $i--) {
        if ($keep -notcontains $i) { $pres.Slides.Item($i).Delete() }
    }
    $pres.SaveAs($Out, 24)
    if (-not (Test-Path -LiteralPath $RenderDir)) { New-Item -ItemType Directory -Path $RenderDir -Force | Out-Null }
    $w = 1920
    $h = [int]([math]::Round($w * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth))
    for ($n = 1; $n -le $pres.Slides.Count; $n++) {
        $pres.Slides.Item($n).Export((Join-Path $RenderDir ("slide-{0:d2}.png" -f $n)), "PNG", $w, $h)
    }
    Write-Output ("built " + $pres.Slides.Count + " slides -> " + $Out)
} finally {
    $pres.Close()
    $ppt.Quit()
}
