"""Fill a real .pptx template with new content by driving PowerPoint (COM).

Why COM instead of editing the OOXML package by hand: PowerPoint produces a
guaranteed-valid file and keeps the template design (master, layouts, media)
intact, which hand-rebuilding the zip repeatedly broke.

Shape addressing: shape indexes returned by the COM object model do NOT match
the order of <p:sp> elements in the slide XML, so addressing shapes by index
is unreliable. Shapes are therefore matched by their CURRENT TEXT: each slide
entry supplies {current text: new text}. A shape whose text equals (or is a
prefix of) a key is rewritten; nothing else is touched.

Input : input/xauat-content-text.json  -> {"slides":[{"src":N,"map":{old:new}}]}
Output: output/nylon66-xauat.pptx  +  render/slide-NN.png for visual checks
"""
import json
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PS1 = ROOT / 'src' / 'build_with_powerpoint.ps1'

SCRIPT = r'''
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
'''

PS1.write_text(SCRIPT, encoding='utf-8')
r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(PS1)],
                   capture_output=True, text=True, encoding='utf-8', errors='replace')
print(r.stdout.strip() or r.stderr.strip())
raise SystemExit(r.returncode)
