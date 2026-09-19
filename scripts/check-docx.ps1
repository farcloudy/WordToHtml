<#
  Dump what Word itself sees in a generated .docx, as JSON.

  Why this exists: it is the hardest verification available here. It does not look
  at our code at all, only at Word's own judgement about style names, East Asian
  fonts, font sizes, line spacing rules, spacing before/after, first-line indent
  in character units, page-number fields, tracked changes, comments and
  per-paragraph underline.

  IMPORTANT: keep this file pure ASCII.
  Windows PowerShell 5.1 reads a BOM-less .ps1 as ANSI (GBK on this machine), so
  non-ASCII comments get misdecoded into bytes that break the parser. Node/Vite
  read .mjs/.ts as UTF-8 and are unaffected, so keep the Chinese prose in those.
  Verified the hard way on 2026-09-12.

  Paths are passed in and deliberately go through an ASCII temp path, to avoid
  cmd.exe -> powershell.exe mangling a non-ASCII repo path.
#>
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [Parameter(Mandatory = $true)][string]$OutPath
)

$ErrorActionPreference = 'Stop'

$word = New-Object -ComObject Word.Application
$word.Visible = 0
$word.DisplayAlerts = 0

$dump = [ordered]@{}
$doc = $null

try {
  $doc = $word.Documents.Open($Path, $false, $true)

  $dump.fileName = $doc.Name
  $dump.sectionCount = [int]$doc.Sections.Count
  $dump.revisionCount = [int]$doc.Revisions.Count
  $dump.commentCount = [int]$doc.Comments.Count
  $dump.pageCount = [int]$doc.ComputeStatistics(2)

  $sections = @()
  for ($s = 1; $s -le $doc.Sections.Count; $s++) {
    $sec = $doc.Sections.Item($s)
    $foot = $sec.Footers.Item(1)
    $types = @()
    if ($foot.Range.Fields.Count -gt 0) {
      $types = @($foot.Range.Fields | ForEach-Object { [int]$_.Type })
    }
    $ps = $sec.PageSetup
    $restart = 0
    try { $restart = [int]$sec.PageNumbers.RestartNumberingAtSection } catch { $restart = -1 }
    $sections += [ordered]@{
      index                 = $s
      startingNumber        = [int]$sec.PageNumbers.StartingNumber
      restartNumbering      = $restart
      footerText            = $foot.Range.Text
      footerFieldTypes      = $types
      linkedToPrevious      = [int]$foot.LinkToPrevious
      orientation           = [int]$ps.Orientation
      # Document grid: 0 = none, 1 = lines and chars, 2 = lines only. CharsLine /
      # LinesPage are what Word's page-setup dialog shows for this grid.
      layoutMode            = [int]$ps.LayoutMode
      charsLine             = [int]$ps.CharsLine
      linesPage             = [int]$ps.LinesPage
      pageWidth             = [math]::Round([double]$ps.PageWidth, 2)
      pageHeight            = [math]::Round([double]$ps.PageHeight, 2)
      topMargin             = [math]::Round([double]$ps.TopMargin, 2)
      bottomMargin          = [math]::Round([double]$ps.BottomMargin, 2)
      leftMargin            = [math]::Round([double]$ps.LeftMargin, 2)
      rightMargin           = [math]::Round([double]$ps.RightMargin, 2)
      footerDistance        = [math]::Round([double]$ps.FooterDistance, 2)
    }
  }
  $dump.sections = $sections

  # Paragraph styles only (Type = 1). Not filtered by InUse: the node side picks
  # the styles it wants by name, which is one less judgement call in here.
  #
  # Deliberately NOT reading ParagraphFormat.FirstLineIndent: Word throws when the
  # indent is expressed in character units, which is exactly what we set (2 chars)
  # on body and headings. CharacterUnitFirstLineIndent reports it correctly.
  $styles = @()
  foreach ($st in $doc.Styles) {
    if ([int]$st.Type -ne 1) { continue }
    $pf = $st.ParagraphFormat
    $charIndent = -1
    try { $charIndent = [double]$pf.CharacterUnitFirstLineIndent } catch { $charIndent = -1 }
    $styles += [ordered]@{
      name                         = $st.NameLocal
      fontFarEast                  = $st.Font.NameFarEast
      fontAscii                    = $st.Font.NameAscii
      fontSize                     = [double]$st.Font.Size
      bold                         = [int]$st.Font.Bold
      alignment                    = [int]$pf.Alignment
      lineSpacingRule              = [int]$pf.LineSpacingRule
      lineSpacing                  = [math]::Round([double]$pf.LineSpacing, 2)
      spaceBefore                  = [math]::Round([double]$pf.SpaceBefore, 2)
      spaceAfter                   = [math]::Round([double]$pf.SpaceAfter, 2)
      characterUnitFirstLineIndent = $charIndent
    }
  }
  $dump.styles = $styles

  $paragraphs = @()
  $i = 0
  foreach ($p in $doc.Paragraphs) {
    $r = $p.Range
    # Underline of the text only, excluding the trailing paragraph mark: the mark
    # carries its own formatting, so a fully underlined paragraph would otherwise
    # read back as wdUndefined (9999999, mixed). -1 because the mark is 1 char.
    # A paragraph that is only the mark has no text at all -> report 0 (none).
    $underline = 0
    if (($r.End - $r.Start) -gt 1) {
      $inner = $doc.Range($r.Start, $r.End - 1)
      $underline = [int]$inner.Font.Underline
    }
    # wdWithInTable = 12. Word's Paragraphs collection also walks the paragraphs
    # inside table cells; the node side has to exclude them or its "paragraph
    # order and text" section gets scrambled by table content.
    $inTable = 0
    try { if ([int]$r.Information(12) -ne 0) { $inTable = 1 } } catch { $inTable = 0 }
    $paragraphs += [ordered]@{
      index     = $i
      style     = $p.Style.NameLocal
      text      = $r.Text.TrimEnd([char]13, [char]7, [char]10)
      bold      = [int]$r.Font.Bold
      underline = $underline
      alignment = [int]$p.Alignment
      inTable   = $inTable
    }
    $i++
  }
  $dump.paragraphs = $paragraphs

  # Tables. Two object-model facts drive the shape below:
  #   * Word 16.0 exposes Row.AllowBreakAcrossPages (not Row.CantSplit), so the
  #     dumped cantSplit is derived as its inverse.
  #   * Table.PreferredWidth reports wdUndefined (9999999) as soon as a row has a
  #     horizontally merged cell, so the effective table width is summed from the
  #     cells of the first row instead (a merged cell reports the full width).
  $tables = @()
  $ti = 0
  foreach ($tbl in $doc.Tables) {
    $ti++
    $rowCount = -1
    try { $rowCount = [int]$tbl.Rows.Count } catch { $rowCount = -1 }
    $colCount = -1
    try { $colCount = [int]$tbl.Columns.Count } catch { $colCount = -1 }

    $prefType = -1
    $prefWidth = -1
    try { $prefType = [int]$tbl.PreferredWidthType } catch { $prefType = -1 }
    try { $prefWidth = [math]::Round([double]$tbl.PreferredWidth, 2) } catch { $prefWidth = -1 }

    $widthPoints = -1
    try {
      $sum = 0.0
      foreach ($cell in $tbl.Rows.Item(1).Cells) { $sum += [double]$cell.Width }
      $widthPoints = [math]::Round($sum, 2)
    } catch { $widthPoints = -1 }

    $tableBorders = [ordered]@{}
    foreach ($side in @(@('top', -1), @('left', -2), @('bottom', -3), @('right', -4))) {
      $lineStyle = -1
      $lineWidth = -1
      try {
        $b = $tbl.Borders.Item([int]$side[1])
        $lineStyle = [int]$b.LineStyle
        $lineWidth = [math]::Round([double]$b.LineWidth, 2)
      } catch {
        $lineStyle = -1
        $lineWidth = -1
      }
      $tableBorders[$side[0]] = [ordered]@{ lineStyle = $lineStyle; lineWidth = $lineWidth }
    }

    $rows = @()
    $ri = 0
    try {
      foreach ($row in $tbl.Rows) {
        $ri++
        $columnIndices = @()
        try {
          $columnIndices = @($row.Cells | ForEach-Object { [int]$_.ColumnIndex })
        } catch { $columnIndices = @() }

        $cells = @()
        $ci = 0
        foreach ($cell in $row.Cells) {
          $cellAlign = -1
          try { $cellAlign = [int]$cell.Range.ParagraphFormat.Alignment } catch { $cellAlign = -1 }
          # Paragraph style name of the cell's first paragraph (W4b-2: cells may use a
          # style other than the list-item one). Empty when unavailable.
          $cellStyle = ''
          try { $cellStyle = $cell.Range.Paragraphs.Item(1).Style.NameLocal } catch { $cellStyle = '' }
          $vAlign = -1
          try { $vAlign = [int]$cell.VerticalAlignment } catch { $vAlign = -1 }
          $span = -1
          if ($columnIndices.Count -gt $ci) {
            $next = $colCount + 1
            if ($ci -lt $columnIndices.Count - 1) { $next = $columnIndices[$ci + 1] }
            if ($next -gt $columnIndices[$ci]) { $span = $next - $columnIndices[$ci] }
          }
          $cellBorders = [ordered]@{}
          foreach ($side in @(@('top', -1), @('left', -2), @('bottom', -3), @('right', -4))) {
            $ls = -1
            try { $ls = [int]$cell.Borders.Item([int]$side[1]).LineStyle } catch { $ls = -1 }
            $cellBorders[$side[0]] = $ls
          }
          $cells += [ordered]@{
            text              = $cell.Range.Text.TrimEnd([char]13, [char]7, [char]10)
            style             = $cellStyle
            alignment         = $cellAlign
            verticalAlignment = $vAlign
            columnSpan        = $span
            lineStyles        = $cellBorders
          }
          $ci++
        }

        $cantSplit = -1
        try {
          if ([bool]$row.AllowBreakAcrossPages) { $cantSplit = 0 } else { $cantSplit = 1 }
        } catch { $cantSplit = -1 }
        $heightRule = -1
        try { $heightRule = [int]$row.HeightRule } catch { $heightRule = -1 }
        $height = -1
        try { $height = [math]::Round([double]$row.Height, 2) } catch { $height = -1 }
        $pwt = -1
        try { $pwt = [int]$row.PreferredWidthType } catch { $pwt = -1 }
        $pw = -1
        try { $pw = [math]::Round([double]$row.PreferredWidth, 2) } catch { $pw = -1 }

        $rows += [ordered]@{
          index              = $ri
          cellCount          = [int]$row.Cells.Count
          cantSplit          = $cantSplit
          heightRule         = $heightRule
          height             = $height
          preferredWidthType = $pwt
          preferredWidth     = $pw
          cells              = $cells
        }
      }
    } catch {
      # Rows enumeration can throw when vertical merges are present. The current
      # design only ever merges horizontally (gridSpan), so this is a guard.
      $rowCount = -1
    }

    $tables += [ordered]@{
      index              = $ti
      rowCount           = $rowCount
      columnCount        = $colCount
      preferredWidthType = $prefType
      preferredWidth     = $prefWidth
      widthPoints        = $widthPoints
      borders            = $tableBorders
      rows               = $rows
    }
  }
  $dump.tables = $tables

  # wdRevisionInsert = 1, wdRevisionDelete = 2
  $revisions = @()
  foreach ($rev in $doc.Revisions) {
    $revisions += [ordered]@{
      type   = [int]$rev.Type
      author = $rev.Author
      text   = $rev.Range.Text
    }
  }
  $dump.revisions = $revisions

  $comments = @()
  foreach ($c in $doc.Comments) {
    $comments += [ordered]@{
      index  = [int]$c.Index
      author = $c.Author
      text   = $c.Range.Text
      scope  = $c.Scope.Text
    }
  }
  $dump.comments = $comments

  $doc.Close($false)
}
catch {
  Write-Output ('[ERROR] ' + $_.Exception.Message)
  Write-Output $_.InvocationInfo.PositionMessage
  throw
}
finally {
  # Releasing both objects and forcing a GC is what actually makes WINWORD.EXE exit.
  # Without it Word leaves a hidden instance behind holding the file, and the next
  # run hits EBUSY. Verified the hard way on 2026-09-12.
  if ($doc) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
  if ($word) {
    $word.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  }
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

$json = $dump | ConvertTo-Json -Depth 8
[System.IO.File]::WriteAllText($OutPath, $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Output ('wrote ' + $OutPath)
