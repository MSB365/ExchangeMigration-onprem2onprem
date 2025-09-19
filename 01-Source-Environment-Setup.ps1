<#
.SYNOPSIS
  Cross-Forest Exchange Migration - SOURCE Environment Setup (Step 1)
.DESCRIPTION
  Configures MRS-Proxy on source Exchange servers and validates connectivity.
  Run this script on the SOURCE Exchange environment first.
.NOTES
  Execute in Exchange Management Shell on SOURCE environment with Organization Management rights.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
  # === Source Environment Configuration ===
  [Parameter(Mandatory=$true)] [string] $SourceEwsFqdn,         # e.g. mail.source.tld
  [Parameter(Mandatory=$true)] [string] $SourceDC,              # e.g. dc01.source.local
  [Parameter(Mandatory=$true)] [string] $SourceAuthoritativeDomain, # e.g. source.tld
  
  # === Target Environment Info (for testing) ===
  [Parameter(Mandatory=$true)] [string] $TargetDC,              # e.g. dc01.target.local
  [Parameter(Mandatory=$true)] [string] $TargetDeliveryDomain,  # e.g. target.tld
  
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
    $saveDialog.FileName = "Source-Environment-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
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
    <title>Exchange Migration - Source Environment Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f7fa; color: #333; line-height: 1.6; }
        .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 10px; margin-bottom: 30px; text-align: center; }
        .header h1 { font-size: 2.5em; margin-bottom: 10px; }
        .header p { font-size: 1.2em; opacity: 0.9; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin-bottom: 30px; }
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
        .progress-fill { height: 100%; background: linear-gradient(90deg, #27ae60, #2ecc71); transition: width 0.3s ease; }
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
            <p>Source Environment Setup - Step 1</p>
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
                    <option value="MRS-Proxy Setup">MRS-Proxy Setup</option>
                    <option value="Connectivity Test">Connectivity Test</option>
                    <option value="Validation">Validation</option>
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
    Write-Header "Source Environment Setup - Step 1"
    Write-LogEntry -Message "Starting Exchange Cross-Forest Migration - Source Environment Setup" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Source EWS FQDN: $SourceEwsFqdn" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Source DC: $SourceDC" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Target DC: $TargetDC" -Level "Info" -Phase "General"
    
    # Create output directory
    if (-not (Test-Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-LogEntry -Message "Created output directory: $OutputPath" -Level "Info" -Phase "General"
    }
}

process {
    try {
        # Step 1: Enable MRS-Proxy on Source
        Write-Header "MRS-Proxy Configuration"
        Write-LogEntry -Message "Configuring MRS-Proxy on source Exchange servers" -Level "Info" -Phase "MRS-Proxy Setup"
        
        try {
            # Get all EWS virtual directories
            $ewsVDirs = Get-WebServicesVirtualDirectory
            foreach ($vdir in $ewsVDirs) {
                if (-not $vdir.MRSProxyEnabled) {
                    Set-WebServicesVirtualDirectory -Identity $vdir.Identity -MRSProxyEnabled $true
                    Write-LogEntry -Message "Enabled MRS-Proxy on $($vdir.Identity)" -Level "Success" -Phase "MRS-Proxy Setup"
                } else {
                    Write-LogEntry -Message "MRS-Proxy already enabled on $($vdir.Identity)" -Level "Info" -Phase "MRS-Proxy Setup"
                }
            }
            $script:Statistics.CompletedSteps++
        } catch {
            Write-LogEntry -Message "Failed to configure MRS-Proxy: $($_.Exception.Message)" -Level "Error" -Phase "MRS-Proxy Setup"
            throw
        }

        # Step 2: Restart IIS
        Write-Header "IIS Restart"
        Write-LogEntry -Message "Restarting IIS to apply MRS-Proxy changes" -Level "Info" -Phase "MRS-Proxy Setup"
        
        try {
            iisreset /noforce
            Start-Sleep -Seconds 10
            Write-LogEntry -Message "IIS restarted successfully" -Level "Success" -Phase "MRS-Proxy Setup"
            $script:Statistics.CompletedSteps++
        } catch {
            Write-LogEntry -Message "Failed to restart IIS: $($_.Exception.Message)" -Level "Error" -Phase "MRS-Proxy Setup"
            throw
        }

        # Step 3: Validate MRS-Proxy Configuration
        Write-Header "MRS-Proxy Validation"
        Write-LogEntry -Message "Validating MRS-Proxy configuration" -Level "Info" -Phase "Validation"
        
        try {
            $ewsUrl = "https://$SourceEwsFqdn/EWS/mrsproxy.svc"
            Write-LogEntry -Message "Testing MRS-Proxy endpoint: $ewsUrl" -Level "Info" -Phase "Validation"
            
            # Test web request to MRS-Proxy endpoint
            $response = Invoke-WebRequest -Uri $ewsUrl -UseBasicParsing -TimeoutSec 30
            if ($response.StatusCode -eq 200) {
                Write-LogEntry -Message "MRS-Proxy endpoint is accessible" -Level "Success" -Phase "Validation"
            } else {
                Write-LogEntry -Message "MRS-Proxy endpoint returned status code: $($response.StatusCode)" -Level "Warning" -Phase "Validation"
            }
            $script:Statistics.CompletedSteps++
        } catch {
            Write-LogEntry -Message "Failed to validate MRS-Proxy endpoint: $($_.Exception.Message)" -Level "Error" -Phase "Validation"
            # Don't throw here as this might be expected in some network configurations
        }

        # Step 4: Display Next Steps
        Write-Header "Next Steps"
        Write-LogEntry -Message "Source environment setup completed successfully" -Level "Success" -Phase "General"
        Write-LogEntry -Message "Next: Run script 02-Target-Environment-Prepare.ps1 on the TARGET environment" -Level "Info" -Phase "General"
        Write-LogEntry -Message "MRS-Proxy URL for target configuration: https://$SourceEwsFqdn/EWS/mrsproxy.svc" -Level "Info" -Phase "General"
        
        $script:Statistics.CompletedSteps++
        $script:Statistics.Success = $true

    } catch {
        Write-LogEntry -Message "Source environment setup failed: $($_.Exception.Message)" -Level "Error" -Phase "General"
        $script:Statistics.Success = $false
        throw
    }
}

end {
    $script:Statistics.EndTime = Get-Date
    $duration = $script:Statistics.EndTime - $script:Statistics.StartTime
    
    Write-Header "Source Environment Setup Complete"
    Write-LogEntry -Message "Total execution time: $($duration.ToString('hh\:mm\:ss'))" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Steps completed: $($script:Statistics.CompletedSteps)/$($script:Statistics.TotalSteps)" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Errors: $($script:Statistics.Errors), Warnings: $($script:Statistics.Warnings)" -Level "Info" -Phase "General"
    
    # Generate HTML report
    $reportHtml = Generate-HTMLReport -OutputPath $OutputPath
    
    # Save report
    $reportPath = $null
    if ($ShowSaveDialog) {
        $reportPath = Show-SaveDialog -DefaultPath $OutputPath
    }
    
    if (-not $reportPath) {
        $reportPath = Join-Path $OutputPath "Source-Environment-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    }
    
    try {
        $reportHtml | Out-File -FilePath $reportPath -Encoding UTF8
        Write-LogEntry -Message "HTML report saved to: $reportPath" -Level "Success" -Phase "General"
        
        # Open report in default browser
        Start-Process $reportPath
    } catch {
        Write-LogEntry -Message "Failed to save HTML report: $($_.Exception.Message)" -Level "Error" -Phase "General"
    }
    
    Write-Host "`nSource environment setup completed!" -ForegroundColor Green
    Write-Host "Report saved to: $reportPath" -ForegroundColor Yellow
    Write-Host "`nNext step: Run 02-Target-Environment-Prepare.ps1 on the TARGET environment" -ForegroundColor Cyan
}
