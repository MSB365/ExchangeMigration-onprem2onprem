<#
.SYNOPSIS
  Cross-Forest Exchange Migration - TARGET Environment Cleanup (Step 5)
.DESCRIPTION
  Cleans up completed migration batches, move requests, and optionally removes endpoints.
  Run this script on the TARGET Exchange environment after Step 4.
.NOTES
  Execute in Exchange Management Shell on TARGET environment with Organization Management rights.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
  # === Cleanup Configuration ===
  [Parameter()] [string] $BatchPrefix = "CF-Batch",
  [Parameter()] [string] $EndpointName = "CF-Source-EWS",
  [Parameter()] [string] $OrgRelationshipName = "SourceOrg",
  [Parameter()] [switch] $RemoveEndpoint,
  [Parameter()] [switch] $RemoveOrgRelationship,
  [Parameter()] [switch] $Force,
  
  # === Reporting ===
  [Parameter()] [string] $OutputPath = ".\MigrationReports",
  [Parameter()] [switch] $ShowSaveDialog = $true
)

# Initialize logging and statistics
$script:LogEntries = @()
$script:Statistics = @{
    StartTime = Get-Date
    EndTime = $null
    TotalSteps = 5
    CompletedSteps = 0
    BatchesRemoved = 0
    MoveRequestsRemoved = 0
    EndpointsRemoved = 0
    OrgRelationshipsRemoved = 0
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
    $saveDialog.FileName = "Target-Cleanup-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
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
    <title>Exchange Migration - Target Environment Cleanup Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f7fa; color: #333; line-height: 1.6; }
        .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
        .header { background: linear-gradient(135deg, #e74c3c 0%, #c0392b 100%); color: white; padding: 30px; border-radius: 10px; margin-bottom: 30px; text-align: center; }
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
        .progress-fill { height: 100%; background: linear-gradient(90deg, #e74c3c, #c0392b); transition: width 0.3s ease; }
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
            <p>Target Environment Cleanup - Step 5</p>
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
                <div class="stat-number success">$($script:Statistics.BatchesRemoved)</div>
                <div class="stat-label">Batches Removed</div>
            </div>
            <div class="stat-card">
                <div class="stat-number success">$($script:Statistics.MoveRequestsRemoved)</div>
                <div class="stat-label">Move Requests Removed</div>
            </div>
            <div class="stat-card">
                <div class="stat-number success">$($script:Statistics.EndpointsRemoved)</div>
                <div class="stat-label">Endpoints Removed</div>
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
                    <option value="Move Request Cleanup">Move Request Cleanup</option>
                    <option value="Batch Cleanup">Batch Cleanup</option>
                    <option value="Endpoint Cleanup">Endpoint Cleanup</option>
                    <option value="Organization Relationship Cleanup">Organization Relationship Cleanup</option>
                    <option value="Final Verification">Final Verification</option>
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
    Write-Header "Target Environment Cleanup - Step 5"
    Write-LogEntry -Message "Starting Exchange Cross-Forest Migration - Target Environment Cleanup" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Batch Prefix: $BatchPrefix" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Remove Endpoint: $RemoveEndpoint" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Remove Org Relationship: $RemoveOrgRelationship" -Level "Info" -Phase "General"
    
    # Create output directory
    if (-not (Test-Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-LogEntry -Message "Created output directory: $OutputPath" -Level "Info" -Phase "General"
    }
}

process {
    try {
        # Step 1: Clean up completed move requests
        Write-Header "Move Request Cleanup"
        Write-LogEntry -Message "Cleaning up completed move requests" -Level "Info" -Phase "Move Request Cleanup"
        
        $completedMoveRequests = Get-MoveRequest | Where-Object { $_.Status -eq "Completed" }
        Write-LogEntry -Message "Found $($completedMoveRequests.Count) completed move requests" -Level "Info" -Phase "Move Request Cleanup"
        
        foreach ($moveRequest in $completedMoveRequests) {
            try {
                Write-LogEntry -Message "Removing completed move request: $($moveRequest.DisplayName)" -Level "Info" -Phase "Move Request Cleanup"
                Remove-MoveRequest -Identity $moveRequest.Identity -Confirm:$false
                Write-LogEntry -Message "Successfully removed move request: $($moveRequest.DisplayName)" -Level "Success" -Phase "Move Request Cleanup"
                $script:Statistics.MoveRequestsRemoved++
            } catch {
                Write-LogEntry -Message "Failed to remove move request $($moveRequest.DisplayName): $($_.Exception.Message)" -Level "Error" -Phase "Move Request Cleanup"
            }
        }
        $script:Statistics.CompletedSteps++

        # Step 2: Clean up migration batches
        Write-Header "Migration Batch Cleanup"
        Write-LogEntry -Message "Cleaning up completed migration batches" -Level "Info" -Phase "Batch Cleanup"
        
        $batches = Get-MigrationBatch | Where-Object { $_.Name -like "$BatchPrefix*" }
        $completedBatches = $batches | Where-Object { $_.Status -in @("Completed", "CompletedWithErrors") }
        
        Write-LogEntry -Message "Found $($completedBatches.Count) completed batches out of $($batches.Count) total batches" -Level "Info" -Phase "Batch Cleanup"
        
        foreach ($batch in $completedBatches) {
            try {
                Write-LogEntry -Message "Removing completed batch: $($batch.Name)" -Level "Info" -Phase "Batch Cleanup"
                Remove-MigrationBatch -Identity $batch.Identity -Confirm:$false
                Write-LogEntry -Message "Successfully removed batch: $($batch.Name)" -Level "Success" -Phase "Batch Cleanup"
                $script:Statistics.BatchesRemoved++
            } catch {
                Write-LogEntry -Message "Failed to remove batch $($batch.Name): $($_.Exception.Message)" -Level "Error" -Phase "Batch Cleanup"
            }
        }
        
        # Check for remaining batches
        $remainingBatches = Get-MigrationBatch | Where-Object { $_.Name -like "$BatchPrefix*" }
        if ($remainingBatches.Count -gt 0) {
            Write-LogEntry -Message "Warning: $($remainingBatches.Count) batches still remain (not completed or in use)" -Level "Warning" -Phase "Batch Cleanup"
            foreach ($batch in $remainingBatches) {
                Write-LogEntry -Message "Remaining batch: $($batch.Name) - Status: $($batch.Status)" -Level "Warning" -Phase "Batch Cleanup"
            }
        }
        $script:Statistics.CompletedSteps++

        # Step 3: Remove migration endpoint (optional)
        if ($RemoveEndpoint) {
            Write-Header "Migration Endpoint Cleanup"
            Write-LogEntry -Message "Removing migration endpoint: $EndpointName" -Level "Info" -Phase "Endpoint Cleanup"
            
            $endpoint = Get-MigrationEndpoint -Identity $EndpointName -ErrorAction SilentlyContinue
            if ($endpoint) {
                try {
                    Remove-MigrationEndpoint -Identity $EndpointName -Confirm:$false
                    Write-LogEntry -Message "Successfully removed migration endpoint: $EndpointName" -Level "Success" -Phase "Endpoint Cleanup"
                    $script:Statistics.EndpointsRemoved++
                } catch {
                    Write-LogEntry -Message "Failed to remove migration endpoint: $($_.Exception.Message)" -Level "Error" -Phase "Endpoint Cleanup"
                }
            } else {
                Write-LogEntry -Message "Migration endpoint '$EndpointName' not found" -Level "Warning" -Phase "Endpoint Cleanup"
            }
        } else {
            Write-LogEntry -Message "Skipping migration endpoint removal (not requested)" -Level "Info" -Phase "Endpoint Cleanup"
        }
        $script:Statistics.CompletedSteps++

        # Step 4: Remove organization relationship (optional)
        if ($RemoveOrgRelationship) {
            Write-Header "Organization Relationship Cleanup"
            Write-LogEntry -Message "Removing organization relationship: $OrgRelationshipName" -Level "Info" -Phase "Organization Relationship Cleanup"
            
            $orgRelationship = Get-OrganizationRelationship -Identity $OrgRelationshipName -ErrorAction SilentlyContinue
            if ($orgRelationship) {
                try {
                    Remove-OrganizationRelationship -Identity $OrgRelationshipName -Confirm:$false
                    Write-LogEntry -Message "Successfully removed organization relationship: $OrgRelationshipName" -Level "Success" -Phase "Organization Relationship Cleanup"
                    $script:Statistics.OrgRelationshipsRemoved++
                } catch {
                    Write-LogEntry -Message "Failed to remove organization relationship: $($_.Exception.Message)" -Level "Error" -Phase "Organization Relationship Cleanup"
                }
            } else {
                Write-LogEntry -Message "Organization relationship '$OrgRelationshipName' not found" -Level "Warning" -Phase "Organization Relationship Cleanup"
            }
        } else {
            Write-LogEntry -Message "Skipping organization relationship removal (not requested)" -Level "Info" -Phase "Organization Relationship Cleanup"
        }
        $script:Statistics.CompletedSteps++

        # Step 5: Final verification
        Write-Header "Final Verification"
        Write-LogEntry -Message "Performing final verification of cleanup" -Level "Info" -Phase "Final Verification"
        
        $remainingMoveRequests = Get-MoveRequest | Where-Object { $_.Status -eq "Completed" }
        $remainingBatches = Get-MigrationBatch | Where-Object { $_.Name -like "$BatchPrefix*" }
        
        Write-LogEntry -Message "Remaining completed move requests: $($remainingMoveRequests.Count)" -Level "Info" -Phase "Final Verification"
        Write-LogEntry -Message "Remaining migration batches: $($remainingBatches.Count)" -Level "Info" -Phase "Final Verification"
        
        if ($remainingMoveRequests.Count -eq 0 -and $remainingBatches.Count -eq 0) {
            Write-LogEntry -Message "Cleanup completed successfully - no migration artifacts remaining" -Level "Success" -Phase "Final Verification"
        } else {
            Write-LogEntry -Message "Some migration artifacts still remain - manual cleanup may be required" -Level "Warning" -Phase "Final Verification"
        }
        $script:Statistics.CompletedSteps++

        # Display completion message
        Write-Header "Cleanup Summary"
        Write-LogEntry -Message "Migration cleanup completed successfully" -Level "Success" -Phase "General"
        Write-LogEntry -Message "Next: Optionally run 06-Source-Environment-Teardown.ps1 on the SOURCE environment" -Level "Info" -Phase "General"
        Write-LogEntry -Message "Migration project is now complete!" -Level "Success" -Phase "General"
        
        $script:Statistics.Success = $true

    } catch {
        Write-LogEntry -Message "Cleanup process failed: $($_.Exception.Message)" -Level "Error" -Phase "General"
        $script:Statistics.Success = $false
        throw
    }
}

end {
    $script:Statistics.EndTime = Get-Date
    $duration = $script:Statistics.EndTime - $script:Statistics.StartTime
    
    Write-Header "Target Environment Cleanup Complete"
    Write-LogEntry -Message "Total execution time: $($duration.ToString('hh\:mm\:ss'))" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Steps completed: $($script:Statistics.CompletedSteps)/$($script:Statistics.TotalSteps)" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Move requests removed: $($script:Statistics.MoveRequestsRemoved), Batches removed: $($script:Statistics.BatchesRemoved)" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Endpoints removed: $($script:Statistics.EndpointsRemoved), Org relationships removed: $($script:Statistics.OrgRelationshipsRemoved)" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Errors: $($script:Statistics.Errors), Warnings: $($script:Statistics.Warnings)" -Level "Info" -Phase "General"
    
    # Generate HTML report
    $reportHtml = Generate-HTMLReport -OutputPath $OutputPath
    
    # Save report
    $reportPath = $null
    if ($ShowSaveDialog) {
        $reportPath = Show-SaveDialog -DefaultPath $OutputPath
    }
    
    if (-not $reportPath) {
        $reportPath = Join-Path $OutputPath "Target-Cleanup-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    }
    
    try {
        $reportHtml | Out-File -FilePath $reportPath -Encoding UTF8
        Write-LogEntry -Message "HTML report saved to: $reportPath" -Level "Success" -Phase "General"
        
        # Open report in default browser
        Start-Process $reportPath
    } catch {
        Write-LogEntry -Message "Failed to save HTML report: $($_.Exception.Message)" -Level "Error" -Phase "General"
    }
    
    Write-Host "`nCleanup process completed!" -ForegroundColor Green
    Write-Host "Report saved to: $reportPath" -ForegroundColor Yellow
    Write-Host "`nOptional: Run 06-Source-Environment-Teardown.ps1 on the SOURCE environment to complete teardown" -ForegroundColor Cyan
    Write-Host "Migration project is now complete!" -ForegroundColor Green
}
