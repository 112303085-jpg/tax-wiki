# 由 deck.py render / check-render 呼叫：依繪圖指令 JSON 以本機 PowerPoint 產生 .pptx，
# 並檢查每個文字框是否超出（結果寫入 -Report 的 JSON）；給 -PngDir 時把每張投影片匯出成圖片。
param([string]$Spec, [string]$Out, [string]$Report, [string]$PngDir = "")
$ErrorActionPreference = 'Stop'
$app = $null; $pres = $null
trap {
  Write-Output ("ERR: " + $_.Exception.Message + " | line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.InvocationInfo.Line.Trim())
  if ($pres) { try { $pres.Close() } catch {} }
  if ($app) { try { $app.Quit() } catch {} }
  exit 1
}

$deckSpec = [IO.File]::ReadAllText($Spec, [Text.Encoding]::UTF8) | ConvertFrom-Json  # 注意：PowerShell 變數不分大小寫，不可命名為 $spec
$F = $deckSpec.font
function C([string]$hex) { [Convert]::ToInt32($hex.Substring(0,2),16) + 256*[Convert]::ToInt32($hex.Substring(2,2),16) + 65536*[Convert]::ToInt32($hex.Substring(4,2),16) }
function Style($r, $size, [string]$color, $bold) {
  $r.Font.Name = $F; $r.Font.NameFarEast = $F; $r.Font.Size = [single]$size
  $r.Font.Color.RGB = (C $color); $r.Font.Bold = $(if ($bold) { -1 } else { 0 })
}
$ALIGN = @{ 'l' = 1; 'c' = 2; 'r' = 3 }
$ANCHOR = @{ 't' = 1; 'm' = 3; 'b' = 4 }
$SHAPE = @{ 'rect' = 1; 'round' = 5; 'oval' = 9 }
$issues = New-Object System.Collections.ArrayList

function Issue($no, [string]$what) { [void]$issues.Add(@{ no = [int]$no; issue = $what }) }

function NoMargin($tf) { $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0 }

$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Add(0)
$pres.PageSetup.SlideWidth = 960; $pres.PageSetup.SlideHeight = 540

foreach ($sl in $deckSpec.slides) {
  $s = $pres.Slides.Add($pres.Slides.Count + 1, 12)
  $s.FollowMasterBackground = 0; $s.Background.Fill.Solid(); $s.Background.Fill.ForeColor.RGB = (C 'FFFFFF')
  foreach ($x in $sl.shapes) {
    switch ($x.k) {
      'box' {
        $sh = $s.Shapes.AddShape($SHAPE[[string]$x.shape], [single]$x.x, [single]$x.y, [single]$x.w, [single]$x.h)
        $sh.Fill.Solid(); $sh.Fill.ForeColor.RGB = (C $x.fill)
        if ($x.line) { $sh.Line.Visible = -1; $sh.Line.ForeColor.RGB = (C $x.line); $sh.Line.Weight = [single]$x.lw } else { $sh.Line.Visible = 0 }
        if ($x.shadow) {
          $sh.Shadow.Visible = -1; $sh.Shadow.OffsetX = [single]5; $sh.Shadow.OffsetY = [single]5
          $sh.Shadow.Blur = [single]0; $sh.Shadow.ForeColor.RGB = 0; $sh.Shadow.Transparency = [single]0
        } else { $sh.Shadow.Visible = 0 }
        if ($x.shape -eq 'round') { $sh.Adjustments.Item(1) = [single]0.5 }
        if ($x.rot) { $sh.Rotation = [single]$x.rot }
        if ($x.text) {
          $tf = $sh.TextFrame; NoMargin $tf; $tf.WordWrap = 0; $tf.VerticalAnchor = 3
          $r = $tf.TextRange; $r.Text = [string]$x.text; Style $r $x.size $x.color $x.bold; $r.ParagraphFormat.Alignment = 2
          if ($r.BoundWidth -gt $sh.Width + 2) { Issue $sl.no ("標籤文字超出框：" + [string]$x.text) }
        }
      }
      'text' {
        $sh = $s.Shapes.AddTextbox(1, [single]$x.x, [single]$x.y, [single]$x.w, [single]$x.h)
        $tf = $sh.TextFrame; NoMargin $tf; $tf.AutoSize = 0; $tf.WordWrap = $(if ($x.wrap) { -1 } else { 0 })
        $r = $tf.TextRange; $r.Text = [string]$x.text; Style $r $x.size $x.color $x.bold
        $r.ParagraphFormat.Alignment = $ALIGN[[string]$x.align]
        $tf.AutoSize = 0; $sh.Top = [single]$x.y; $sh.Height = [single]$x.h; $sh.Width = [single]$x.w; $tf.VerticalAnchor = $ANCHOR[[string]$x.anchor]
        if ($x.name) { $sh.Name = [string]$x.name }
        $short = ([string]$x.text); if ($short.Length -gt 16) { $short = $short.Substring(0, 16) + '…' }
        if ($x.wrap) { if ($r.BoundHeight -gt $x.h + 3) { Issue $sl.no ("文字超出框（高）：" + $short) } }
        else { if ($r.BoundWidth -gt $x.w + 3) { Issue $sl.no ("文字超出框（寬）：" + $short) } }
      }
      'bullets' {
        $sh = $s.Shapes.AddTextbox(1, [single]$x.x, [single]$x.y, [single]$x.w, [single]$x.h)
        $tf = $sh.TextFrame; NoMargin $tf; $tf.AutoSize = 0; $tf.WordWrap = -1
        $r = $tf.TextRange; $r.Text = (@($x.items) -join "`r"); Style $r $x.size '111111' $x.bold
        $r.ParagraphFormat.Bullet.Visible = -1; $r.ParagraphFormat.Bullet.Character = 8226
        $r.ParagraphFormat.SpaceAfter = [single]6
        $tf.Ruler.Levels.Item(1).FirstMargin = [single]0; $tf.Ruler.Levels.Item(1).LeftMargin = [single]18
        $tf.AutoSize = 0; $sh.Top = [single]$x.y; $sh.Height = [single]$x.h
        if ($r.BoundHeight -gt $x.h + 3) { Issue $sl.no ("條列超出框：" + (@($x.items)[0])) }
      }
      'pic' {
        [void]$s.Shapes.AddPicture([string]$x.path, 0, -1, [single]$x.x, [single]$x.y, [single]$x.w, [single]$x.h)
      }
      'table' {
        $rows = @($x.rows); $nR = $rows.Count; $nC = @($rows[0]).Count
        $w = 0; foreach ($cw in $x.colw) { $w += [double]$cw }
        $sh = $s.Shapes.AddTable($nR, $nC, [single]$x.x, [single]$x.y, [single]$w, [single]($x.rowh * $nR))
        $t = $sh.Table
        for ($c = 1; $c -le $nC; $c++) { $t.Columns.Item($c).Width = [single]$x.colw[$c-1] }
        for ($i = 1; $i -le $nR; $i++) {
          $t.Rows.Item($i).Height = [single]$x.rowh
          for ($c = 1; $c -le $nC; $c++) {
            $cell = $t.Cell($i, $c); $ctf = $cell.Shape.TextFrame
            $ctf.MarginLeft = [single]9; $ctf.MarginRight = [single]9; $ctf.MarginTop = [single]3; $ctf.MarginBottom = [single]3; $ctf.VerticalAnchor = 3
            $tr = $ctf.TextRange; $tr.Text = [string](@($rows[$i-1])[$c-1])
            $cell.Shape.Fill.Solid()
            if ($i -eq 1) { $cell.Shape.Fill.ForeColor.RGB = (C '111111'); Style $tr $x.size 'FFFFFF' $true }
            else {
              $cell.Shape.Fill.ForeColor.RGB = (C 'FFFFFF')
              $em = @($x.em_rows) -contains ($i - 1)
              Style $tr $x.size $(if ($em) { '2D5BFF' } else { '111111' }) $em
            }
            if (@($x.right_cols) -contains ($c - 1)) { $tr.ParagraphFormat.Alignment = 3 }
            foreach ($b in 1..4) { $bd = $cell.Borders.Item($b); $bd.Visible = -1; $bd.ForeColor.RGB = 0; $bd.Weight = [single]1.5 }
          }
        }
        if ($x.maxh -and ($sh.Height -gt [double]$x.maxh + 2)) { Issue $sl.no ("表格太高：" + [math]::Round($sh.Height) + "pt，上限 " + [math]::Round($x.maxh) + "pt") }
      }
    }
  }
  if ($sl.notes) { $s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = [string]$sl.notes }
}

$pres.SaveAs($Out, 24)
if ($PngDir) {
  New-Item -ItemType Directory -Force $PngDir | Out-Null
  Get-ChildItem $PngDir -Filter 'pptx-*.png' | Remove-Item
  # PowerPoint 的 Export 無法寫入含中文的路徑：先匯出到英文暫存資料夾再搬移
  $tmp = Join-Path ([IO.Path]::GetTempPath()) ('deckpng-' + [guid]::NewGuid().ToString('N'))
  New-Item -ItemType Directory -Force $tmp | Out-Null
  foreach ($sl in $pres.Slides) { $sl.Export((Join-Path $tmp ('pptx-{0:D2}.png' -f $sl.SlideIndex)), 'PNG', 1280, 720) }
  Get-ChildItem $tmp -Filter 'pptx-*.png' | Move-Item -Destination $PngDir
  Remove-Item $tmp -Recurse -Force
  if ((Get-ChildItem $PngDir -Filter 'pptx-*.png').Count -ne $pres.Slides.Count) { Issue 0 "PowerPoint 匯出圖片數量不符" }
}
$pres.Close(); $app.Quit()
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
$json = ConvertTo-Json -InputObject @($issues) -Depth 4
[IO.File]::WriteAllText($Report, $json, (New-Object Text.UTF8Encoding $false))
Write-Output ("OK slides=" + @($deckSpec.slides).Count + " issues=" + $issues.Count)
