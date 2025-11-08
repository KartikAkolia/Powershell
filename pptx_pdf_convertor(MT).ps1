# Limit to 4-6 parallel PowerPoint instances (safer for COM objects)
$maxThreads = 4  # Conservative value that works reliably
# OR use: $maxThreads = [Math]::Min(6, ((Get-CimInstance Win32_ComputerSystem).NumberOfLogicalProcessors - 2))

Write-Host "Using $maxThreads parallel threads for conversion`n"

# Get ONLY PPTX files
$files = Get-ChildItem -Path "*.pptx" -Recurse | Where-Object { $_.Extension -eq ".pptx" }

Write-Host "Found $($files.Count) PPTX files to convert`n"

$completedCount = 0
$totalCount = $files.Count

# Process files in parallel
$files | ForEach-Object -Parallel {
    $pptx = $_.FullName
    $pdf = $pptx -replace '\.pptx$', '.pdf'
    
    # Skip if PDF already exists
    if (Test-Path $pdf) {
        Write-Host "⏭ Skipped (already exists): $($_.Name)" -ForegroundColor Yellow
        return
    }
    
    $powerpoint = $null
    $presentation = $null
    
    try {
        # Add retry logic for COM initialization
        $retryCount = 0
        $maxRetries = 3
        
        while ($retryCount -lt $maxRetries) {
            try {
                $powerpoint = New-Object -ComObject PowerPoint.Application
                break
            }
            catch {
                $retryCount++
                if ($retryCount -eq $maxRetries) { throw }
                Start-Sleep -Milliseconds 500
            }
        }
        
        $presentation = $powerpoint.Presentations.Open($pptx, $true, $true, $false)
        $presentation.SaveAs($pdf, 32)
        $presentation.Close()
        
        $syncCount = $using:completedCount
        $syncTotal = $using:totalCount
        Write-Host "✓ [$syncCount/$syncTotal] Converted: $($_.Name)" -ForegroundColor Green
        
    }
    catch {
        Write-Host "✗ Failed: $($_.Name) - Error: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        # Proper cleanup
        if ($presentation) {
            try { $presentation.Close() } catch {}
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
        }
        if ($powerpoint) {
            try { $powerpoint.Quit() } catch {}
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
    
} -ThrottleLimit $maxThreads

Write-Host "`n✅ All conversions complete! Processed $($files.Count) files." -ForegroundColor Green