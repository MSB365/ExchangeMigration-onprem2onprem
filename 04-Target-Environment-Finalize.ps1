<#
.SYNOPSIS
  Cross-Forest Exchange Migration - TARGET Environment Finalization (Step 4)
.DESCRIPTION
  Finalizes migration batches, resumes suspended move requests, and monitors completion.
  Run this script on the TARGET Exchange environment after Step 3.
.NOTES
  Execute in Exchange Management Shell on TARGET environment with Organization Management rights.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
  # === Migration Configuration ===
  [Parameter()] [string] $BatchPrefix = "CF-Batch",
  [Parameter()] [ValidateSet("Finalize", "Resume", "Monitor")] [string] $Action = "Finalize",
  [Parameter()] [switch] $AutoComplete,
  
  # === Reporting ===
  [Parameter()] [string] $OutputPath = ".\MigrationReports",
  [Parameter()] [switch] $ShowSaveDialog = $true
)

# Initialize logging and statistics
$script:LogEntries = @()
$script:Statistics = @{
    StartTime = Get-Date
    EndTime = $null
    TotalSteps = 4
    CompletedSteps = 0
    BatchesProcessed = 0
    MoveRequestsResumed = 0
    CompletedMigrations = 0
    Errors = 0
    Warnings = 0
    Success = $false
}

function Write-LogEntry {
    param(
        [Parameter(Mandatory=$true)] [string] $Message,
        [Parameter()] [ValidateSet("Info", "Warning", "Error", "Success")] [string] $Level = "Info",
        [Parameter()] [string] $Phase = "General"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = [PSCustomObject]@{
        Timestamp = $timestamp
        Phase = $Phase
        Level = $Level
        Message = $Message
    }
    
    $script:LogEntries += $entry
    
    # Update statistics
    switch ($Level) {
        "Error" { $script:Statistics.Errors++ }
        "Warning" { $script:Statistics.Warnings++ }
    }
    
    # Console output with colors
    $color = switch ($Level) {
        "Info" { "White" }
        "Warning" { "Yellow" }
        "Error" { "Red" }
        "Success" { "Green" }
    }
    
    Write-Host "[$timestamp] [$Phase] $Message" -ForegroundColor $color
}

function Write-Header($text) {
    Write-Host "`n=== $text ===" -ForegroundColor Cyan
    Write-LogEntry -Message "Starting phase: $text" -Level "Info" -Phase $text
}

function Show-SaveDialog {
    param([string] $DefaultPath)
    
    Add-Type -AssemblyName System.Windows.Forms
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "HTML files (*.html)|*.html|All files (*.*)|*.*"
    $saveDialog.Title = "Save Migration Report"
    $saveDialog.FileName = "Target-Finalize-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    $saveDialog.InitialDirectory = $DefaultPath
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $saveDialog.FileName
    }
    return $null
}

function Generate-HTMLReport {
    param([string] $OutputPath)
    
    $reportHtml = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exchange Migration - Target Environment Finalization Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f7fa; color: #333; line-height: 1.6; }
        .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
        .header { background: linear-gradient(135deg, #8e44ad 0%, #9b59b6 100%); color: white; padding: 30px; border-radius: 10px; margin-bottom: 30px; text-align: center; }
        .header h1 { font-size: 2.5em; margin-bottom: 10px; }
        .header p { font-size: 1.2em; opacity: 0.9; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .stat-card { background: white; padding: 25px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center; }
        .stat-number { font-size: 2.5em; font-weight: bold; margin-bottom: 10px; }
        .stat-label { color: #666; font-size: 1.1em; }
        .success { color: #27ae60; }
        .warning { color: #f39c12; }
        .error { color: #e74c3c; }
        .info { color: #3498db; }
        .log-section { background: white; border-radius: 10px; padding: 30px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .log-controls { margin-bottom: 20px; display: flex; gap: 15px; flex-wrap: wrap; align-items: center; }
        .log-controls input, .log-controls select { padding: 10px; border: 1px solid #ddd; border-radius: 5px; font-size: 14px; }
        .log-controls input[type="text"] { flex: 1; min-width: 200px; }
        .log-entry { padding: 12px; margin-bottom: 8px; border-radius: 5px; border-left: 4px solid; }
        .log-entry.info { background: #f8f9ff; border-color: #3498db; }
        .log-entry.warning { background: #fffbf0; border-color: #f39c12; }
        .log-entry.error { background: #fdf2f2; border-color: #e74c3c; }
        .log-entry.success { background: #f0f9f4; border-color: #27ae60; }
        .log-timestamp { font-weight: bold; color: #666; }
        .log-phase { background: #ecf0f1; padding: 2px 8px; border-radius: 3px; font-size: 0.9em; margin: 0 5px; }
        .log-message { margin-top: 5px; }
        .progress-bar { width: 100%; height: 20px; background: #ecf0f1; border-radius: 10px; overflow: hidden; margin: 20px 0; }
        .progress-fill { height: 100%; background: linear-gradient(90deg, #8e44ad, #9b59b6); transition: width 0.3s ease; }
        .hidden { display: none; }
        @media (max-width: 768px) {
            .stats-grid { grid-template-columns: 1fr; }
            .log-controls { flex-direction: column; align-items: stretch; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Exchange Migration Report</h1>
            <p>Target Environment Finalization - Step 4</p>
            <p>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        </div>

        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-number success">$($script:Statistics.CompletedSteps)</div>
                <div class="stat-label">Completed Steps</div>
            </div>
            <div class="stat-card">
                <div class="stat-number info">$($script:Statistics.TotalSteps)</div>
                <div class="stat-label">Total Steps</div>
            </div>
            <div class="stat-card">
                <div class="stat-number success">$($script:Statistics.BatchesProcessed)</div>
                <div class="stat-label">Batches Processed</div>
            </div>
            <div class="stat-card">
                <div class="stat-number success">$($script:Statistics.MoveRequestsResumed)</div>
                <div class="stat-label">Move Requests Resumed</div>
            </div>
            <div class="stat-card">
                <div class="stat-number success">$($script:Statistics.CompletedMigrations)</div>
                <div class="stat-label">Completed Migrations</div>
            </div>
            <div class="stat-card">
                <div class="stat-number error">$($script:Statistics.Errors)</div>
                <div class="stat-label">Errors</div>
            </div>
        </div>

        <div class="progress-bar">
            <div class="progress-fill" style="width: $(($script:Statistics.CompletedSteps / $script:Statistics.TotalSteps) * 100)%"></div>
        </div>

        <div class="log-section">
            <h2>Execution Log</h2>
            <div class="log-controls">
                <input type="text" id="searchInput" placeholder="Search logs..." onkeyup="filterLogs()">
                <select id="phaseFilter" onchange="filterLogs()">
                    <option value="">All Phases</option>
                    <option value="General">General</option>
                    <option value="Batch Finalization">Batch Finalization</option>
                    <option value="Move Request Resume">Move Request Resume</option>
                    <option value="Migration Monitoring">Migration Monitoring</option>
                    <option value="Status Check">Status Check</option>
                </select>
                <select id="levelFilter" onchange="filterLogs()">
                    <option value="">All Levels</option>
                    <option value="Info">Info</option>
                    <option value="Warning">Warning</option>
                    <option value="Error">Error</option>
                    <option value="Success">Success</option>
                </select>
            </div>
            <div id="logEntries">
"@

    foreach ($entry in $script:LogEntries) {
        $reportHtml += @"
                <div class="log-entry $($entry.Level.ToLower())" data-phase="$($entry.Phase)" data-level="$($entry.Level)">
                    <span class="log-timestamp">$($entry.Timestamp)</span>
                    <span class="log-phase">$($entry.Phase)</span>
                    <div class="log-message">$($entry.Message)</div>
                </div>
"@
    }

    $reportHtml += @"
            </div>
        </div>
    </div>

    <script>
        function filterLogs() {
            const searchTerm = document.getElementById('searchInput').value.toLowerCase();
            const phaseFilter = document.getElementById('phaseFilter').value;
            const levelFilter = document.getElementById('levelFilter').value;
            const entries = document.querySelectorAll('.log-entry');

            entries.forEach(entry => {
                const message = entry.textContent.toLowerCase();
                const phase = entry.getAttribute('data-phase');
                const level = entry.getAttribute('data-level');

                const matchesSearch = message.includes(searchTerm);
                const matchesPhase = !phaseFilter || phase === phaseFilter;
                const matchesLevel = !levelFilter || level === levelFilter;

                if (matchesSearch && matchesPhase && matchesLevel) {
                    entry.classList.remove('hidden');
                } else {
                    entry.classList.add('hidden');
                }
            });
        }
    </script>
</body>
</html>
"@

    return $reportHtml
}

begin {
    Write-Header "Target Environment Finalization - Step 4"
    Write-LogEntry -Message "Starting Exchange Cross-Forest Migration - Target Environment Finalization" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Action: $Action" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Batch Prefix: $BatchPrefix" -Level "Info" -Phase "General"
    
    # Create output directory
    if (-not (Test-Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-LogEntry -Message "Created output directory: $OutputPath" -Level "Info" -Phase "General"
    }
}

process {
    try {
        # Step 1: Get current migration status
        Write-Header "Migration Status Check"
        Write-LogEntry -Message "Checking current migration status" -Level "Info" -Phase "Status Check"
        
        $batches = Get-MigrationBatch | Where-Object { $_.Name -like "$BatchPrefix*" }
        $moveRequests = Get-MoveRequest
        
        Write-LogEntry -Message "Found $($batches.Count) migration batches" -Level "Info" -Phase "Status Check"
        Write-LogEntry -Message "Found $($moveRequests.Count) move requests" -Level "Info" -Phase "Status Check"
        
        foreach ($batch in $batches) {
            Write-LogEntry -Message "Batch: $($batch.Name), Status: $($batch.Status), Progress: $($batch.PercentageComplete)%" -Level "Info" -Phase "Status Check"
        }
        $script:Statistics.CompletedSteps++

        # Step 2: Handle based on action
        switch ($Action) {
            "Finalize" {
                Write-Header "Batch Finalization"
                Write-LogEntry -Message "Finalizing migration batches" -Level "Info" -Phase "Batch Finalization"
                
                $syncedBatches = $batches | Where-Object { $_.Status -eq "Synced" }
                foreach ($batch in $syncedBatches) {
                    try {
                        Write-LogEntry -Message "Finalizing batch: $($batch.Name)" -Level "Info" -Phase "Batch Finalization"
                        Complete-MigrationBatch -Identity $batch.Identity -Confirm:$false
                        Write-LogEntry -Message "Successfully finalized batch: $($batch.Name)" -Level "Success" -Phase "Batch Finalization"
                        $script:Statistics.BatchesProcessed++
                    } catch {
                        Write-LogEntry -Message "Failed to finalize batch $($batch.Name): $($_.Exception.Message)" -Level "Error" -Phase "Batch Finalization"
                    }
                }
                
                if ($syncedBatches.Count -eq 0) {
                    Write-LogEntry -Message "No batches in 'Synced' status found for finalization" -Level "Warning" -Phase "Batch Finalization"
                }
            }
            
            "Resume" {
                Write-Header "Move Request Resume"
                Write-LogEntry -Message "Resuming suspended move requests" -Level "Info" -Phase "Move Request Resume"
                
                $suspendedRequests = $moveRequests | Where-Object { $_.Status -eq "Suspended" -or $_.Status -eq "AutoSuspended" }
                foreach ($request in $suspendedRequests) {
                    try {
                        Write-LogEntry -Message "Resuming move request: $($request.DisplayName)" -Level "Info" -Phase "Move Request Resume"
                        Resume-MoveRequest -Identity $request.Identity -Confirm:$false
                        Write-LogEntry -Message "Successfully resumed move request: $($request.DisplayName)" -Level "Success" -Phase "Move Request Resume"
                        $script:Statistics.MoveRequestsResumed++
                    } catch {
                        Write-LogEntry -Message "Failed to resume move request $($request.DisplayName): $($_.Exception.Message)" -Level "Error" -Phase "Move Request Resume"
                    }
                }
                
                if ($suspendedRequests.Count -eq 0) {
                    Write-LogEntry -Message "No suspended move requests found" -Level "Info" -Phase "Move Request Resume"
                }
            }
            
            "Monitor" {
                Write-Header "Migration Monitoring"
                Write-LogEntry -Message "Monitoring migration progress" -Level "Info" -Phase "Migration Monitoring"
                
                # Detailed status for each batch
                foreach ($batch in $batches) {
                    $batchStats = Get-MigrationBatchStatistics -Identity $batch.Identity
                    Write-LogEntry -Message "Batch: $($batch.Name) - Total: $($batchStats.TotalCount), Completed: $($batchStats.FinalizedCount), Failed: $($batchStats.FailedCount)" -Level "Info" -Phase "Migration Monitoring"
                }
                
                # Move request statistics
                $mrStats = Get-MoveRequest | Get-MoveRequestStatistics | Group-Object StatusDetail
                foreach ($status in $mrStats) {
                    Write-LogEntry -Message "Move Requests - $($status.Name): $($status.Count)" -Level "Info" -Phase "Migration Monitoring"
                }
            }
        }
        $script:Statistics.CompletedSteps++

        # Step 3: Get updated status
        Write-Header "Updated Status"
        Write-LogEntry -Message "Retrieving updated migration status" -Level "Info" -Phase "Status Check"
        
        $updatedBatches = Get-MigrationBatch | Where-Object { $_.Name -like "$BatchPrefix*" }
        $completedBatches = $updatedBatches | Where-Object { $_.Status -eq "Completed" }
        $completedMoveRequests = Get-MoveRequest | Where-Object { $_.Status -eq "Completed" }
        
        $script:Statistics.CompletedMigrations = $completedMoveRequests.Count
        
        Write-LogEntry -Message "Completed batches: $($completedBatches.Count)/$($updatedBatches.Count)" -Level "Info" -Phase "Status Check"
        Write-LogEntry -Message "Completed move requests: $($completedMoveRequests.Count)" -Level "Info" -Phase "Status Check"
        $script:Statistics.CompletedSteps++

        # Step 4: Display next steps
        Write-Header "Next Steps"
        if ($completedBatches.Count -eq $updatedBatches.Count -and $completedBatches.Count -gt 0) {
            Write-LogEntry -Message "All migration batches completed successfully!" -Level "Success" -Phase "General"
            Write-LogEntry -Message "Next: Run script 05-Target-Environment-Cleanup.ps1 to clean up migration artifacts" -Level "Info" -Phase "General"
        } else {
            Write-LogEntry -Message "Migration still in progress. Monitor status and re-run this script as needed." -Level "Info" -Phase "General"
            Write-LogEntry -Message "Use -Action Monitor to check progress, -Action Resume for suspended requests" -Level "Info" -Phase "General"
        }
        $script:Statistics.CompletedSteps++
        $script:Statistics.Success = $true

    } catch {
        Write-LogEntry -Message "Finalization process failed: $($_.Exception.Message)" -Level "Error" -Phase "General"
        $script:Statistics.Success = $false
        throw
    }
}

end {
    $script:Statistics.EndTime = Get-Date
    $duration = $script:Statistics.EndTime - $script:Statistics.StartTime
    
    Write-Header "Target Environment Finalization Complete"
    Write-LogEntry -Message "Total execution time: $($duration.ToString('hh\:mm\:ss'))" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Steps completed: $($script:Statistics.CompletedSteps)/$($script:Statistics.TotalSteps)" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Batches processed: $($script:Statistics.BatchesProcessed), Move requests resumed: $($script:Statistics.MoveRequestsResumed)" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Completed migrations: $($script:Statistics.CompletedMigrations)" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Errors: $($script:Statistics.Errors), Warnings: $($script:Statistics.Warnings)" -Level "Info" -Phase "General"
    
    # Generate HTML report
    $reportHtml = Generate-HTMLReport -OutputPath $OutputPath
    
    # Save report
    $reportPath = $null
    if ($ShowSaveDialog) {
        $reportPath = Show-SaveDialog -DefaultPath $OutputPath
    }
    
    if (-not $reportPath) {
        $reportPath = Join-Path $OutputPath "Target-Finalize-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    }
    
    try {
        $reportHtml | Out-File -FilePath $reportPath -Encoding UTF8
        Write-LogEntry -Message "HTML report saved to: $reportPath" -Level "Success" -Phase "General"
        
        # Open report in default browser
        Start-Process $reportPath
    } catch {
        Write-LogEntry -Message "Failed to save HTML report: $($_.Exception.Message)" -Level "Error" -Phase "General"
    }
    
    Write-Host "`nFinalization process completed!" -ForegroundColor Green
    Write-Host "Report saved to: $reportPath" -ForegroundColor Yellow
    
    if ($script:Statistics.CompletedMigrations -gt 0) {
        Write-Host "`nNext step: Run 05-Target-Environment-Cleanup.ps1 to clean up migration artifacts" -ForegroundColor Cyan
    } else {
        Write-Host "`nContinue monitoring migration progress and re-run this script as needed" -ForegroundColor Cyan
    }
}
