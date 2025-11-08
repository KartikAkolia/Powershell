# Convert all PPTX files in current directory AND all subdirectories
$powerpoint = New-Object -ComObject PowerPoint.Application

Get-ChildItem -Path "*.pptx" -Recurse | ForEach-Object {
    $pptx = $_.FullName
    $pdf = $_.FullName -replace '\.pptx$', '.pdf'
    
    $presentation = $powerpoint.Presentations.Open($pptx, $true, $true, $false)
    $presentation.SaveAs($pdf, 32)
    $presentation.Close()
    
    Write-Host "Converted: $($_.FullName)"
}

$powerpoint.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint) | Out-Null

Write-Host "`nConversion complete!"