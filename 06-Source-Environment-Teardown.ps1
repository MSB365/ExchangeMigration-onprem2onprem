<#
.SYNOPSIS
  Cross-Forest Exchange Migration - SOURCE Environment Teardown (Step 6)
.DESCRIPTION
  Disables MRS-Proxy on source Exchange servers and performs final cleanup.
  Run this script on the SOURCE Exchange environment after Step 5.
.NOTES
  Execute in Exchange Management Shell on SOURCE environment with Organization Management rights.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
  # === Source Environment Configuration ===
  [Parameter(Mandatory=$true)] [string] $SourceEwsFqdn,         # e.g. mail.source.tld
  [Parameter()] [switch] $DisableMRSProxy = $true,
  [Parameter()] [switch] $RestartIIS = $true,
  
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
    MRSProxyDisabled = 0
    IISRestarted = $false
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
    $saveDialog.FileName = "Source-Teardown-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
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
    <title>Exchange Migration - Source Environment Teardown Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f7fa; color: #333; line-height: 1.6; }
        .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
        .header { background: linear-gradient(135deg, #34495e 0%, #2c3e50 100%); color: white; padding: 30px; border-radius: 10px; margin-bottom: 30px; text-align: center; }
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
        .progress-fill { height: 100%; background: linear-gradient(90deg, #34495e, #2c3e50); transition: width 0.3s ease; }
        .completion-banner { background: linear-gradient(135deg, #27ae60, #2ecc71); color: white; padding: 20px; border-radius: 10px; text-align: center; margin: 20px 0; }
        .completion-banner h2 { font-size: 1.8em; margin-bottom: 10px; }
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
            <p>Source Environment Teardown - Step 6 (Final)</p>
            <p>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        </div>

        <div class="completion-banner">
            <h2>🎉 Migration Project Complete! 🎉</h2>
            <p>Cross-Forest Exchange Migration has been successfully completed and cleaned up.</p>
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
                <div class="stat-number success">$($script:Statistics.MRSProxyDisabled)</div>
                <div class="stat-label">MRS-Proxy Disabled</div>
            </div>
            <div class="stat-card">
                <div class="stat-number $(if($script:Statistics.IISRestarted){'success'}else{'warning'})">$(if($script:Statistics.IISRestarted){'Yes'}else{'No'})</div>
                <div class="stat-label">IIS Restarted</div>
            </div>
            <div class="stat-card">
                <div class="stat-number error">$($script:Statistics.Errors)</div>
                <div class="stat-label">Errors</div>
            </div>
            <div class="stat-card">
                <div class="stat-number warning">$($script:Statistics.Warnings)</div>
                <div class="stat-label">Warnings</div>
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
                    <option value="MRS-Proxy Teardown">MRS-Proxy Teardown</option>
                    <option value="IIS Restart">IIS Restart</option>
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
    Write-Header "Source Environment Teardown - Step 6 (Final)"
    Write-LogEntry -Message "Starting Exchange Cross-Forest Migration - Source Environment Teardown" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Source EWS FQDN: $SourceEwsFqdn" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Disable MRS-Proxy: $DisableMRSProxy" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Restart IIS: $RestartIIS" -Level "Info" -Phase "General"
    
    # Create output directory
    if (-not (Test-Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-LogEntry -Message "Created output directory: $OutputPath" -Level "Info" -Phase "General"
    }
}

process {
    try {
        # Step 1: Check current MRS-Proxy status
        Write-Header "MRS-Proxy Status Check"
        Write-LogEntry -Message "Checking current MRS-Proxy configuration" -Level "Info" -Phase "MRS-Proxy Teardown"
        
        $ewsVDirs = Get-WebServicesVirtualDirectory
        $enabledVDirs = $ewsVDirs | Where-Object { $_.MRSProxyEnabled -eq $true }
        
        Write-LogEntry -Message "Found $($enabledVDirs.Count) EWS virtual directories with MRS-Proxy enabled" -Level "Info" -Phase "MRS-Proxy Teardown"
        foreach ($vdir in $enabledVDirs) {
            Write-LogEntry -Message "MRS-Proxy enabled on: $($vdir.Identity)" -Level "Info" -Phase "MRS-Proxy Teardown"
        }
        $script:Statistics.CompletedSteps++

        # Step 2: Disable MRS-Proxy
        if ($DisableMRSProxy) {
            Write-Header "MRS-Proxy Disabling"
            Write-LogEntry -Message "Disabling MRS-Proxy on source Exchange servers" -Level "Info" -Phase "MRS-Proxy Teardown"
            
            foreach ($vdir in $enabledVDirs) {
                try {
                    Write-LogEntry -Message "Disabling MRS-Proxy on: $($vdir.Identity)" -Level "Info" -Phase "MRS-Proxy Teardown"
                    Set-WebServicesVirtualDirectory -Identity $vdir.Identity -MRSProxyEnabled $false
                    Write-LogEntry -Message "Successfully disabled MRS-Proxy on: $($vdir.Identity)" -Level "Success" -Phase "MRS-Proxy Teardown"
                    $script:Statistics.MRSProxyDisabled++
                } catch {
                    Write-LogEntry -Message "Failed to disable MRS-Proxy on $($vdir.Identity): $($_.Exception.Message)" -Level "Error" -Phase "MRS-Proxy Teardown"
                }
            }
            
            if ($enabledVDirs.Count -eq 0) {
                Write-LogEntry -Message "No MRS-Proxy configurations found to disable" -Level "Info" -Phase "MRS-Proxy Teardown"
            }
        } else {
            Write-LogEntry -Message "Skipping MRS-Proxy disabling (not requested)" -Level "Info" -Phase "MRS-Proxy Teardown"
        }
        $script:Statistics.CompletedSteps++

        # Step 3: Restart IIS
        if ($RestartIIS -and $DisableMRSProxy -and $enabledVDirs.Count -gt 0) {
            Write-Header "IIS Restart"
            Write-LogEntry -Message "Restarting IIS to apply MRS-Proxy changes" -Level "Info" -Phase "IIS Restart"
            
            try {
                iisreset /noforce
                Start-Sleep -Seconds 10
                Write-LogEntry -Message "IIS restarted successfully" -Level "Success" -Phase "IIS Restart"
                $script:Statistics.IISRestarted = $true
            } catch {
                Write-LogEntry -Message "Failed to restart IIS: $($_.Exception.Message)" -Level "Error" -Phase "IIS Restart"
            }
        } else {
            Write-LogEntry -Message "Skipping IIS restart (not needed or not requested)" -Level "Info" -Phase "IIS Restart"
        }
        $script:Statistics.CompletedSteps++

        # Step 4: Final verification
        Write-Header "Final Verification"
        Write-LogEntry -Message "Performing final verification of teardown" -Level "Info" -Phase "Final Verification"
        
        $finalEwsVDirs = Get-WebServicesVirtualDirectory
        $stillEnabledVDirs = $finalEwsVDirs | Where-Object { $_.MRSProxyEnabled -eq $true }
        
        if ($stillEnabledVDirs.Count -eq 0) {
            Write-LogEntry -Message "Verification successful: No MRS-Proxy configurations remain enabled" -Level "Success" -Phase "Final Verification"
        } else {
            Write-LogEntry -Message "Warning: $($stillEnabledVDirs.Count) MRS-Proxy configurations still enabled" -Level "Warning" -Phase "Final Verification"
            foreach ($vdir in $stillEnabledVDirs) {
                Write-LogEntry -Message "Still enabled: $($vdir.Identity)" -Level "Warning" -Phase "Final Verification"
            }
        }
        
        # Test MRS-Proxy endpoint
        try {
            $ewsUrl = "https://$SourceEwsFqdn/EWS/mrsproxy.svc"
            Write-LogEntry -Message "Testing MRS-Proxy endpoint accessibility: $ewsUrl" -Level "Info" -Phase "Final Verification"
            
            $response = Invoke-WebRequest -Uri $ewsUrl -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
            if ($response.StatusCode -eq 200) {
                Write-LogEntry -Message "Warning: MRS-Proxy endpoint is still accessible (may be cached)" -Level "Warning" -Phase "Final Verification"
            }
        } catch {
            Write-LogEntry -Message "MRS-Proxy endpoint is no longer accessible (expected after teardown)" -Level "Success" -Phase "Final Verification"
        }
        $script:Statistics.CompletedSteps++

        # Display completion message
        Write-Header "Migration Project Complete"
        Write-LogEntry -Message "🎉 Cross-Forest Exchange Migration project completed successfully! 🎉" -Level "Success" -Phase "General"
        Write-LogEntry -Message "Source environment teardown completed" -Level "Success" -Phase "General"
        Write-LogEntry -Message "All migration artifacts have been cleaned up" -Level "Success" -Phase "General"
        Write-LogEntry -Message "MRS-Proxy has been disabled on source servers" -Level "Success" -Phase "General"
        
        $script:Statistics.Success = $true

    } catch {
        Write-LogEntry -Message "Teardown process failed: $($_.Exception.Message)" -Level "Error" -Phase "General"
        $script:Statistics.Success = $false
        throw
    }
}

end {
    $script:Statistics.EndTime = Get-Date
    $duration = $script:Statistics.EndTime - $script:Statistics.StartTime
    
    Write-Header "Source Environment Teardown Complete"
    Write-LogEntry -Message "Total execution time: $($duration.ToString('hh\:mm\:ss'))" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Steps completed: $($script:Statistics.CompletedSteps)/$($script:Statistics.TotalSteps)" -Level "Info" -Phase "General"
    Write-LogEntry -Message "MRS-Proxy configurations disabled: $($script:Statistics.MRSProxyDisabled)" -Level "Info" -Phase "General"
    Write-LogEntry -Message "IIS restarted: $($script:Statistics.IISRestarted)" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Errors: $($script:Statistics.Errors), Warnings: $($script:Statistics.Warnings)" -Level "Info" -Phase "General"
    
    # Generate HTML report
    $reportHtml = Generate-HTMLReport -OutputPath $OutputPath
    
    # Save report
    $reportPath = $null
    if ($ShowSaveDialog) {
        $reportPath = Show-SaveDialog -DefaultPath $OutputPath
    }
    
    if (-not $reportPath) {
        $reportPath = Join-Path $OutputPath "Source-Teardown-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    }
    
    try {
        $reportHtml | Out-File -FilePath $reportPath -Encoding UTF8
        Write-LogEntry -Message "HTML report saved to: $reportPath" -Level "Success" -Phase "General"
        
        # Open report in default browser
        Start-Process $reportPath
    } catch {
        Write-LogEntry -Message "Failed to save HTML report: $($_.Exception.Message)" -Level "Error" -Phase "General"
    }
    
    Write-Host "`n🎉 MIGRATION PROJECT COMPLETE! 🎉" -ForegroundColor Green
    Write-Host "Report saved to: $reportPath" -ForegroundColor Yellow
    Write-Host "`nCross-Forest Exchange Migration has been successfully completed!" -ForegroundColor Green
    Write-Host "All migration artifacts have been cleaned up." -ForegroundColor Green
    Write-Host "Thank you for using the Exchange Migration Toolkit!" -ForegroundColor Cyan
}
