# Limit to 4-6 parallel instances (safer for COM objects)
$maxThreads = 4  # Conservative value that works reliably
# OR use: $maxThreads = [Math]::Min(6, ((Get-CimInstance Win32_ComputerSystem).NumberOfLogicalProcessors - 2))

Write-Host "Using $maxThreads parallel threads for conversion`n"

# Get PPTX and DOCX files
$pptxFiles = Get-ChildItem -Path "*.pptx" -Recurse | Where-Object { $_.Extension -eq ".pptx" }
$docxFiles = Get-ChildItem -Path "*.docx" -Recurse | Where-Object { $_.Extension -eq ".docx" }
$allFiles = $pptxFiles + $docxFiles

Write-Host "Found $($pptxFiles.Count) PPTX files and $($docxFiles.Count) DOCX files to convert"
Write-Host "Total: $($allFiles.Count) files`n"

$completedCount = 0
$totalCount = $allFiles.Count

# Process files in parallel
$allFiles | ForEach-Object -Parallel {
    $sourceFile = $_.FullName
    $fileExtension = $_.Extension.ToLower()
    $pdf = $sourceFile -replace '\.(pptx|docx)$', '.pdf'
    
    # Skip if PDF already exists
    if (Test-Path $pdf) {
        Write-Host "⏭ Skipped (already exists): $($_.Name)" -ForegroundColor Yellow
        return
    }
    
    $app = $null
    $doc = $null
    
    try {
        # Add retry logic for COM initialization
        $retryCount = 0
        $maxRetries = 3
        
        while ($retryCount -lt $maxRetries) {
            try {
                if ($fileExtension -eq ".pptx") {
                    $app = New-Object -ComObject PowerPoint.Application
                } elseif ($fileExtension -eq ".docx") {
                    $app = New-Object -ComObject Word.Application
                    $app.Visible = $false
                }
                break
            }
            catch {
                $retryCount++
                if ($retryCount -eq $maxRetries) { throw }
                Start-Sleep -Milliseconds 500
            }
        }
        
        # Open and convert based on file type
        if ($fileExtension -eq ".pptx") {
            $doc = $app.Presentations.Open($sourceFile, $true, $true, $false)
            $doc.SaveAs($pdf, 32)  # 32 = ppSaveAsPDF
        } elseif ($fileExtension -eq ".docx") {
            $doc = $app.Documents.Open($sourceFile)
            $doc.SaveAs([ref]$pdf, [ref]17)  # 17 = wdFormatPDF
        }
        
        $doc.Close()
        
        $syncCount = $using:completedCount
        $syncTotal = $using:totalCount
        Write-Host "✓ [$syncCount/$syncTotal] Converted: $($_.Name)" -ForegroundColor Green
        
    }
    catch {
        Write-Host "✗ Failed: $($_.Name) - Error: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        # Proper cleanup
        if ($doc) {
            try { $doc.Close() } catch {}
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        }
        if ($app) {
            try { $app.Quit() } catch {}
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
    
} -ThrottleLimit $maxThreads

Write-Host "`n✅ All conversions complete! Processed $($allFiles.Count) files." -ForegroundColor Green
