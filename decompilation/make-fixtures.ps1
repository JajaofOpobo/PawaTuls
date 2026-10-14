# make-fixtures.ps1 — Phase 29: generate Word-run fixture documents.
# Run: powershell -NoProfile -ExecutionPolicy Bypass -File .\decompilation\make-fixtures.ps1
$ErrorActionPreference = "Stop"
$root = Get-Location
$fixDir = Join-Path $root "decompilation\fixtures"
New-Item -ItemType Directory -Force -Path $fixDir | Out-Null

$word = New-Object -ComObject Word.Application
try {
  $word.Visible = $false
  $word.DisplayAlerts = 0

  # ── formatting-test.docx ──────────────────────────────────────────────
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $sel.Style = $doc.Styles.Item("Heading 1")
  $sel.TypeText("Formatting Fixture")
  $sel.TypeParagraph()
  $sel.Style = $doc.Styles.Item("Normal")
  $sel.TypeText("Plain body paragraph one.")
  $sel.TypeParagraph()
  $sel.Font.Bold = $true
  $sel.TypeText("Bold lead")
  $sel.Font.Bold = $false
  $sel.TypeText(" and ")
  $sel.Font.Italic = $true
  $sel.TypeText("italic tail")
  $sel.Font.Italic = $false
  $sel.TypeText(" mixed.")
  $sel.TypeParagraph()
  $sel.Font.Size = 20
  $sel.Font.Color = 255
  $sel.TypeText("Big red line.")
  $sel.Font.Size = 11
  $sel.Font.Color = 0
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Alignment = 2  # center
  $sel.TypeText("Centered line")
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Alignment = 3  # right
  $sel.TypeText("Right line")
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Alignment = 3
  $sel.TypeText("justify  me  doubLES  and	tabs")
  $sel.TypeParagraph()
  $sel.ParagraphFormat.Alignment = 0
  $sel.TypeText("lowercase line for UPPERCASE test")
  $sel.TypeParagraph()
  $sel.TypeText("SENTENCE CASE SOURCE. second sentence here! third one?")
  $sel.TypeParagraph()
  $sel.TypeText("First manual line")
  $sel.TypeText([char]11)  # manual line break
  $sel.TypeText("second manual line")
  $sel.TypeParagraph()
  $sel.TypeText("1st 2nd 3rd 4th 21st 200th ordinals line")
  $sel.TypeParagraph()
  $sel.TypeText("--------------------------------------------------")
  $sel.TypeParagraph()
  $sel.TypeText("")
  $sel.TypeParagraph()  # truly empty
  $sel.TypeText("   ")
  $sel.TypeParagraph()  # whitespace-only
  # native numbered list
  $sel.Range.ListFormat.ApplyNumberDefault() | Out-Null
  $sel.TypeText("native one")
  $sel.TypeParagraph()
  $sel.TypeText("native two")
  $sel.TypeParagraph()
  $sel.Range.ListFormat.RemoveNumbers() | Out-Null
  # table with an empty row
  $rng = $sel.Range
  $table = $doc.Tables.Add($rng, 3, 2)
  $table.Cell(1, 1).Range.Text = "A1"
  $table.Cell(1, 2).Range.Text = "B1"
  $table.Cell(2, 1).Range.Text = ""
  $table.Cell(2, 2).Range.Text = ""
  $table.Cell(3, 1).Range.Text = "A3"
  $table.Cell(3, 2).Range.Text = "B3"
  # genuine footnote + endnote (must be left alone by scanner)
  $sel.EndKey(6) | Out-Null  # wdStory
  $sel.TypeParagraph()
  $sel.TypeText("Body text with a genuine footnote")
  $fnRange = $sel.Range
  $fn = $doc.Footnotes.Add($fnRange)
  $fn.Range.Text = "Genuine footnote text."
  $word.ActiveWindow.View.SeekView = 0  # back to main story
  $sel.EndKey(6) | Out-Null
  $sel.TypeText(" and genuine endnote")
  $enRange = $sel.Range
  $en = $doc.Endnotes.Add($enRange)
  $en.Range.Text = "Genuine endnote text."
  $word.ActiveWindow.View.SeekView = 0
  $sel.EndKey(6) | Out-Null
  $sel.TypeText(".")
  # hyperlink + comment + bookmark
  $sel.TypeParagraph()
  $sel.TypeText("See example site ")
  $hlStart = $sel.Range.Start
  $sel.TypeText("link here")
  $hlRange = $doc.Range($hlStart, $sel.Range.Start)
  $doc.Hyperlinks.Add($hlRange, "https://example.com") | Out-Null
  $sel.TypeParagraph()
  $cRange = $sel.Range
  $sel.TypeText("Commented text.")
  $doc.Comments.Add($cRange, "Reviewer note.") | Out-Null
  $bmRange = $sel.Range
  $doc.Bookmarks.Add("FixtureMark", $bmRange) | Out-Null
  # page break + hidden text
  $sel.TypeParagraph()
  $sel.InsertBreak(7) | Out-Null  # wdPageBreak
  $sel.Font.Hidden = $true
  $sel.TypeText("Hidden fixture text.")
  $sel.Font.Hidden = $false
  $sel.TypeParagraph()
  $sel.TypeText("After page break.")
  $doc.SaveAs([string](Join-Path $fixDir "formatting-test.docx"))
  $doc.Close()
  Write-Host "saved formatting-test.docx"

  # ── footnotes-test.docx ───────────────────────────────────────────────
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $sel.Style = $doc.Styles.Item("Normal")
  $sel.TypeText("Treaty history")
  $sel.Font.Superscript = $true
  $sel.TypeText("1")
  $sel.Font.Superscript = $false
  $sel.TypeText(" shows early agreement")
  $sel.Font.Superscript = $true
  $sel.TypeText("12")
  $sel.Font.Superscript = $false
  $sel.TypeText(" and later text [4] with brackets (5) in parens.")
  $sel.TypeParagraph()
  $sel.TypeText("Unmatched marker")
  $sel.Font.Superscript = $true
  $sel.TypeText("7")
  $sel.Font.Superscript = $false
  $sel.TypeText(" has no block.")
  $sel.TypeParagraph()
  $sel.TypeText("________________________________________")
  $sel.TypeParagraph()
  $sel.TypeText("1 See Vienna Convention on the Law of Treaties.")
  $sel.TypeParagraph()
  $sel.TypeText("12 Continued note text for the twelfth marker spanning lines.")
  $sel.TypeParagraph()
  $sel.TypeText("still part of note twelve.")
  $sel.TypeParagraph()
  $sel.TypeText("[4] See British Yearbook of International Law.")
  $sel.TypeParagraph()
  $sel.TypeText("(5) See American Journal of International Law.")
  $sel.TypeParagraph()
  $sel.TypeText("9 Orphan block with no marker in body.")
  $sel.TypeParagraph()
  $sel.TypeText("[2020] 15 NWLR (Pt. 1721) 123 is a citation, not a footnote.")
  $sel.TypeParagraph()
  $doc.SaveAs([string](Join-Path $fixDir "footnotes-test.docx"))
  $doc.Close()
  Write-Host "saved footnotes-test.docx"

  # ── numbering-test.docx ───────────────────────────────────────────────
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $sel.Style = $doc.Styles.Item("Normal")
  $sel.Range.ListFormat.ApplyNumberDefault() | Out-Null
  $sel.TypeText("native alpha")
  $sel.TypeParagraph()
  $sel.TypeText("native beta")
  $sel.TypeParagraph()
  $sel.Range.ListFormat.RemoveNumbers() | Out-Null
  foreach ($t in @("1. manual one", "2. manual two", "2. manual dup", "4. manual gap")) {
    $sel.TypeText($t)
    $sel.TypeParagraph()
  }
  $sel.TypeText("Ordinary 2026 prose with 3.14 decimals, not a list.")
  $sel.TypeParagraph()
  $doc.SaveAs([string](Join-Path $fixDir "numbering-test.docx"))
  $doc.Close()
  Write-Host "saved numbering-test.docx"

  # ── pleadings-test.docx (fixture for later legal re-enable) ───────────
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $sel.Style = $doc.Styles.Item("Normal")
  $sel.TypeText("The Claimant states as follows:")
  $sel.TypeParagraph()
  $sel.TypeText("First averment paragraph.")
  $sel.TypeParagraph()
  $sel.ParagraphFormat.LeftIndent = 28.35
  $sel.TypeText("Sub-particular under first averment.")
  $sel.TypeParagraph()
  $sel.ParagraphFormat.LeftIndent = 0
  $sel.TypeText("Second averment paragraph.")
  $sel.TypeParagraph()
  $doc.SaveAs([string](Join-Path $fixDir "pleadings-test.docx"))
  $doc.Close()
  Write-Host "saved pleadings-test.docx"
}
finally {
  $word.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
Get-ChildItem $fixDir | Select-Object Name, @{n='KB';e={[int]($_.Length/1KB)}}
