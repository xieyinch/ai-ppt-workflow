"""One-command PPT build: fill a real .pptx template and verify the result.

Workflow
--------
1. Read deck.config.json (template, content map, output path).
2. Drive the installed PowerPoint over COM: open the template, rewrite the
   text of the selected slides, delete the unused slides, save.
3. Verify: PowerPoint re-opens the output, template placeholder text must be
   gone, slide count must match, and every slide is exported to PNG.
4. Print a single PASS/FAIL summary.

Run:  python src/build.py
"""
import argparse
import json
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser(description='Fill a real PPTX template through PowerPoint COM.')
parser.add_argument('--config', default=str(ROOT / 'deck.config.json'), help='JSON config path')
args = parser.parse_args()
config_path = Path(args.config).resolve()
CFG = json.loads(config_path.read_text(encoding='utf-8'))
PS1 = ROOT / 'src' / 'fill_deck.ps1'

required = ['template', 'content', 'output', 'render_dir']
missing = [key for key in required if key not in CFG]
if missing:
    raise SystemExit('config missing required keys: ' + ', '.join(missing))

# text that must not survive in the output: leftover template placeholders
PLACEHOLDERS = [
    '请在这里输入你的正文', '单击此处输入标题', '这里是小标题', '输入小标题',
    '在此输入小标题', '这里写最重要的观点', 'THANKS for all', 'Minzu',
]

PS = r'''
param(
    [string]$Template, [string]$Content, [string]$Out, [string]$RenderDir,
    [string]$Root
)
$ErrorActionPreference = 'Stop'
$cfg = Get-Content -LiteralPath $Content -Raw -Encoding UTF8 | ConvertFrom-Json

# Replace text by exact/prefix match. A map value may be an array: repeated
# placeholder shapes consume array items in document order, avoiding duplicate
# content across sibling cards.
function Walk($shape, $map, $seen) {
    if ($shape.Type -eq 6) { foreach ($c in $shape.GroupItems) { Walk $c $map $seen }; return }
    if (-not $shape.HasTextFrame) { return }
    if (-not $shape.TextFrame.HasText) { return }
    $tr = $shape.TextFrame.TextRange
    $txt = $tr.Text
    if ([string]::IsNullOrEmpty($txt)) { return }
    $hit = $null
    foreach ($old in $map.PSObject.Properties.Name) { if ($txt -eq $old) { $hit = $old; break } }
    if ($null -eq $hit) {
        foreach ($old in $map.PSObject.Properties.Name) { if ($txt.StartsWith($old)) { $hit = $old; break } }
    }
    if ($null -ne $hit) {
        if (-not $seen.ContainsKey($hit)) { $seen[$hit] = 0 }
        $idx = [int]$seen[$hit]
        $seen[$hit] = $idx + 1
        $val = $map.$hit
        if ($val -is [System.Array]) {
            if ($idx -ge $val.Count) { return }
            $new = [string]$val[$idx]
        } else {
            $new = [string]$val
        }
        if ($new -ne $txt) {
            $tr.Text = $new
            try { $shape.TextFrame.AutoSize = 2 } catch { }
        }
    }
}

function ReplaceImages($sld, $entries) {
    if ($null -eq $entries) { return }
    foreach ($entry in $entries) {
        $kind = if ($entry.kind) { [string]$entry.kind } else { 'picture' }
        $target = $null
        if ($entry.name) {
            foreach ($shape in $sld.Shapes) {
                if ($shape.Name -eq [string]$entry.name) { $target = $shape; break }
            }
        } else {
            $wantTypes = if ($kind -eq 'group' -or $kind -eq 'icon') { @(6) } else { @(13, 28) }
            $occ = if ($entry.occurrence) { [int]$entry.occurrence } else { 1 }
            $seen = 0
            foreach ($shape in $sld.Shapes) {
                $isPicture = ([int]$shape.Type -eq 13 -or [int]$shape.Type -eq 28)
                if ([int]$shape.Type -eq 14) { try { $isPicture = ($shape.PlaceholderFormat.Type -eq 18) } catch {} }
                $isGroup = ([int]$shape.Type -eq 6)
                if (($kind -eq 'group' -or $kind -eq 'icon') -and $isGroup) {
                    $seen++; if ($seen -eq $occ) { $target = $shape; break }
                } elseif (($kind -ne 'group' -and $kind -ne 'icon') -and $isPicture) {
                    $seen++; if ($seen -eq $occ) { $target = $shape; break }
                }
            }
        }
        if ($null -eq $target) { throw "image target not found on slide $($sld.SlideIndex): $($entry | ConvertTo-Json -Compress)" }
        $file = [string]$entry.file
        if (-not [System.IO.Path]::IsPathRooted($file)) { $file = Join-Path $Root $file }
        if (-not (Test-Path -LiteralPath $file)) { throw "image file not found: $file" }
        if ([int]$target.Type -eq 14) {
            $left = $target.Left; $top = $target.Top; $width = $target.Width; $height = $target.Height
            $target.Delete()
            $new = $sld.Shapes.AddPicture($file, $false, $true, $left, $top, $width, $height)
        } else {
            $left = $target.Left; $top = $target.Top; $width = $target.Width; $height = $target.Height
            $new = $sld.Shapes.AddPicture($file, $false, $true, $left, $top, $width, $height)
            $target.Delete()
        }
    }
}

function UpdateCharts($sld, $entries) {
    if ($null -eq $entries) { return }
    foreach ($entry in $entries) {
        $target = $null
        if ($entry.name) {
            foreach ($shape in $sld.Shapes) { if ($shape.Name -eq [string]$entry.name) { $target = $shape; break } }
        } else {
            $occ = if ($entry.occurrence) { [int]$entry.occurrence } else { 1 }; $seen = 0
            foreach ($shape in $sld.Shapes) {
                try { $has = $shape.HasChart } catch { $has = $false }
                if ($has) { $seen++; if ($seen -eq $occ) { $target = $shape; break } }
            }
        }
        if ($null -eq $target) { throw "chart target not found on slide $($sld.SlideIndex)" }
        $chart = $target.Chart
        $chart.ChartData.Activate()
        $wb = $chart.ChartData.Workbook
        $ws = $wb.Worksheets.Item(1)
        $cats = @($entry.categories)
        $series = @($entry.series.PSObject.Properties)
        $ws.Cells.Clear()
        $ws.Cells.Item(1,1) = 'Category'
        $col = 2
        foreach ($prop in $series) { $ws.Cells.Item(1,$col) = [string]$prop.Name; $col++ }
        for ($i=0; $i -lt $cats.Count; $i++) {
            $row = $i + 2; $ws.Cells.Item($row,1) = [string]$cats[$i]
            $col = 2
            foreach ($prop in $series) { $vals = @($prop.Value); $ws.Cells.Item($row,$col) = $vals[$i]; $col++ }
        }
        $collection = $chart.SeriesCollection()
        $idx = 1
        foreach ($prop in $series) {
            if ($idx -le $collection.Count) { $ser = $collection.Item($idx) } else { $ser = $collection.NewSeries() }
            $ser.Name = [string]$prop.Name
            $ser.XValues = $cats
            $ser.Values = @($prop.Value)
            $idx++
        }
        $wb.Close($true)
    }
}

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($Template, $false, $false, $false)
try {
    foreach ($entry in $cfg.slides) {
        $sld = $pres.Slides.Item([int]$entry.src)
        $seen = @{}
        foreach ($shape in $sld.Shapes) { Walk $shape $entry.map $seen }
        ReplaceImages $sld $entry.images
    }
    $keep = @($cfg.slides | ForEach-Object { [int]$_.src })
    for ($i = $pres.Slides.Count; $i -ge 1; $i--) {
        if ($keep -notcontains $i) { $pres.Slides.Item($i).Delete() }
    }
    $pres.SaveAs($Out, 24)
    $pres.Close()

    # verification pass: reopen the produced file and export every slide
    $chk = $ppt.Presentations.Open($Out, $true, $false, $false)
    if (-not (Test-Path -LiteralPath $RenderDir)) { New-Item -ItemType Directory -Path $RenderDir -Force | Out-Null }
    $w = 1920
    $h = [int]([math]::Round($w * $chk.PageSetup.SlideHeight / $chk.PageSetup.SlideWidth))
    for ($n = 1; $n -le $chk.Slides.Count; $n++) {
        $chk.Slides.Item($n).Export((Join-Path $RenderDir ("slide-{0:d2}.png" -f $n)), "PNG", $w, $h)
    }
    Write-Output ("reopened OK, slides=" + $chk.Slides.Count)
    $chk.Close()
} finally {
    $ppt.Quit()
}
'''

PS1.write_text(PS, encoding='utf-8')

content_path = ROOT / CFG['content']
out_path = ROOT / CFG['output']
render_dir = ROOT / CFG['render_dir']
content = json.loads(content_path.read_text(encoding='utf-8'))
if CFG.get('design', {}).get('auto_layout'):
    plan_path = out_path.with_name(out_path.stem + '-layout-plan.json')
    if not plan_path.exists():
        raise SystemExit('auto_layout requested but layout plan is missing: ' + str(plan_path))
    plan = json.loads(plan_path.read_text(encoding='utf-8'))
    planned = plan.get('slides', [])
    if len(planned) != len(content.get('slides', [])):
        raise SystemExit('layout plan slide count does not match content')
    for entry, page in zip(content['slides'], planned):
        if page.get('template_src') is not None:
            entry['src'] = int(page['template_src'])
template_path = ROOT / CFG['template']
render_dir.mkdir(parents=True, exist_ok=True)
for old_render in render_dir.glob('slide-*.png'):
    old_render.unlink()
chart_template = None
if any(entry.get('charts') for entry in content.get('slides', [])):
    chart_template = ROOT / '.chart-template-working.pptx'
    shutil.copy2(template_path, chart_template)
    cr = subprocess.run([sys.executable, str(ROOT / 'src' / 'update_charts.py'),
                         str(chart_template), str(content_path), '--source'],
                        capture_output=True, text=True, encoding='utf-8', errors='replace')
    if cr.returncode != 0:
        print(cr.stdout.strip())
        print(cr.stderr.strip())
        sys.exit(1)
    template_path = chart_template

steps = []
r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(PS1),
                    '-Template', str(template_path),
                    '-Content', str(content_path),
                    '-Out', str(out_path),
                    '-RenderDir', str(render_dir),
                    '-Root', str(ROOT)],
                   capture_output=True, text=True, encoding='utf-8', errors='replace')
if r.returncode != 0:
    print(r.stdout.strip())
    print(r.stderr.strip())
    sys.exit(1)
steps.append(r.stdout.strip())

# --- structural checks ---
expected = len(content['slides'])
with zipfile.ZipFile(out_path) as z:
    slides = sorted([n for n in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml', n)],
                    key=lambda s: int(re.search(r'(\d+)', s.split('/')[-1]).group(1)))
    joined = ' '.join(z.read(n).decode('utf-8') for n in slides)
left = sorted({p for p in PLACEHOLDERS if p in joined})
rendered = sorted(render_dir.glob('slide-*.png'))

if chart_template and chart_template.exists():
    chart_template.unlink()

ok = (len(slides) == expected) and not left and len(rendered) == expected
print(json.dumps({
    'template': CFG['template'],
    'output': CFG['output'],
    'slides': len(slides),
    'expected_slides': expected,
    'leftover_placeholders': left,
    'rendered_pngs': len(rendered),
    'status': 'PASS' if ok else 'FAIL',
}, ensure_ascii=False, indent=2))
sys.exit(0 if ok else 1)
