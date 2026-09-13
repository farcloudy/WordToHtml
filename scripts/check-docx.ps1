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
    $paragraphs += [ordered]@{
      index     = $i
      style     = $p.Style.NameLocal
      text      = $r.Text.TrimEnd([char]13, [char]7, [char]10)
      bold      = [int]$r.Font.Bold
      underline = $underline
      alignment = [int]$p.Alignment
    }
    $i++
  }
  $dump.paragraphs = $paragraphs

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
