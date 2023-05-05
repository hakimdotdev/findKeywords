$searchStrings = $args
$rootFolder = 'C:\Your\Path\'
$results = @()

$files = Get-ChildItem -Path $rootFolder -Recurse -Include *.doc, *.docx, *.rtf | Select-Object -ExpandProperty FullName

foreach ($file in $files) {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $doc = $word.Documents.Open($file)

  $lineNumber = 0
  foreach ($paragraph in $doc.Paragraphs) {
    $lineNumber++
    $text = $paragraph.Range.Text
    foreach ($searchString in $searchStrings) {
      if ($text -match $searchString) {
        $result = [PSCustomObject]@{
          File = $file
          LineNumber = $lineNumber
          Text = $text
        }
        $results += $result
      }
    }
  }

  $doc.Close()
  $word.Quit()
}

$results | Format-Table -Property File, LineNumber, Text
