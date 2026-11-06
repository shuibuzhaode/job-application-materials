$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ppt = New-Object -ComObject PowerPoint.Application
Get-ChildItem -LiteralPath $dir -Filter *.pptx | ForEach-Object {
    $out = Join-Path $dir ("png_" + $_.BaseName)
    New-Item -ItemType Directory -Force -Path $out | Out-Null
    $pres = $ppt.Presentations.Open($_.FullName, $true, $false, $false)
    $pres.Export($out, "PNG", 1280, 720)
    $pres.Close()
    Write-Output ("exported: " + $_.Name)
}
$ppt.Quit()
