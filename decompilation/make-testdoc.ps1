# make-testdoc.ps1 — Phase 1: create kutools-test.docx baseline + verify KuTools COM add-in loads.
# Run: powershell -NoProfile -ExecutionPolicy Bypass -File .\decompilation\make-testdoc.ps1
$ErrorActionPreference = "Stop"
$out = Join-Path (Get-Location) "decompilation\kutools-test.docx"
if (Test-Path $out) { Remove-Item $out -Force }

$word = New-Object -ComObject Word.Application
try {
  $word.Visible = $false
  $word.DisplayAlerts = 0  # wdAlertsNone
  Write-Host "---COM ADDINS---"
  foreach ($a in $word.COMAddIns) {
    if ($a.ProgId -like "*Kutools*" -or $a.ProgId -like "*Tabsfor*") {
      Write-Host ("{0}  Connect={1}  Desc={2}" -f $a.ProgId, $a.Connect, $a.Description)
    }
  }
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $sel.Style = $doc.Styles.Item("Heading 1")
  $sel.TypeText("Reference Heading for KuTools Inspection")
  $sel.TypeParagraph()
  $sel.Style = $doc.Styles.Item("Normal")
  $sel.TypeText("First body paragraph with plain normal text for baseline comparison.")
  $sel.TypeParagraph()
  $sel.TypeText("Second body paragraph with plain normal text for baseline comparison.")
  $sel.TypeParagraph()
  $sel.TypeText("Third body paragraph with plain normal text for baseline comparison.")
  $sel.TypeParagraph()
  # Mixed formatting paragraph: bold + italic runs
  $sel.TypeText("Mixed ")
  $sel.Font.Bold = $true
  $sel.TypeText("bold")
  $sel.Font.Bold = $false
  $sel.TypeText(" and ")
  $sel.Font.Italic = $true
  $sel.TypeText("italic")
  $sel.Font.Italic = $false
  $sel.TypeText(" words.")
  $sel.TypeParagraph()
  # 2x3 table
  $range = $sel.Range
  $table = $doc.Tables.Add($range, 2, 3)
  $n = 1
  foreach ($row in $table.Rows) {
    foreach ($cell in $row.Cells) {
      $cell.Range.Text = "R$($cell.RowIndex)C$($cell.ColumnIndex)"
      $n++
    }
  }
  $doc.SaveAs([string]$out)
  $doc.Close()
  Write-Host "saved: $out"
}
finally {
  $word.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
