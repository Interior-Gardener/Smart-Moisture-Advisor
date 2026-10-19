$ErrorActionPreference = 'Stop'
$taskDir = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$qaDir = Join-Path $PSScriptRoot 'qa'
New-Item -ItemType Directory -Path $qaDir -Force | Out-Null
$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false
$wordApp.DisplayAlerts = 0
try {
    $docPath = Join-Path $taskDir 'SMA_Major_Project_Report_10_Pages.docx'
    $reportDoc = $wordApp.Documents.Open($docPath, $false, $false)
    $reportDoc.Fields.Update() | Out-Null
    $reportDoc.Repaginate()
    $pages = $reportDoc.ComputeStatistics(2)
    $reportDoc.Save()
    $reportDoc.ExportAsFixedFormat((Join-Path $taskDir 'SMA_Major_Project_Report_10_Pages.pdf'),17)
    $reportDoc.Close(0)
    Write-Output "Word rendered page count: $pages"
} finally {
    $wordApp.Quit()
}
$pptApp = New-Object -ComObject PowerPoint.Application
try {
    $deckPath = Join-Path $taskDir 'SMA_Major_Project_Presentation.pptx'
    $deck = $pptApp.Presentations.Open($deckPath,$false,$false,$false)
    $slideCount = $deck.Slides.Count
    $deck.Export($qaDir,'PNG',1600,900)
    $deck.SaveAs((Join-Path $taskDir 'SMA_Major_Project_Presentation.pdf'),32)
    $deck.Close()
    Write-Output "PowerPoint rendered slide count: $slideCount"
} finally {
    $pptApp.Quit()
}
