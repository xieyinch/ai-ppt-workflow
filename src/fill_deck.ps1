
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
