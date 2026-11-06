$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ppt = New-Object -ComObject PowerPoint.Application
Get-ChildItem -LiteralPath $dir -Filter *.pptx | ForEach-Object {
    $pres = $ppt.Presentations.Open($_.FullName, $true, $false, $false)
    $pres.SaveAs([System.IO.Path]::ChangeExtension($_.FullName, ".pdf"), 32)
    $pres.Close()
    Write-Output ("pdf: " + $_.Name)
}
$ppt.Quit()
$word = New-Object -ComObject Word.Application
$word.Visible = $false
Get-ChildItem -LiteralPath $dir -Filter *.docx | ForEach-Object {
    $d = $word.Documents.Open($_.FullName)
    $d.SaveAs([ref][System.IO.Path]::ChangeExtension($_.FullName, ".pdf"), [ref]17)
    $d.Close($false)
    Write-Output ("pdf: " + $_.Name)
}
$word.Quit()
