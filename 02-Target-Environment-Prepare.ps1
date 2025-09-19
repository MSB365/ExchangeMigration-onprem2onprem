<#
.SYNOPSIS
  Cross-Forest Exchange Migration - TARGET Environment Preparation (Step 2)
.DESCRIPTION
  Prepares MailUser objects in target forest using Prepare-MoveRequest.ps1.
  Run this script on the TARGET Exchange environment after Step 1.
.NOTES
  Execute in Exchange Management Shell on TARGET environment with Organization Management rights.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
  # === Source Environment Info ===
  [Parameter(Mandatory=$true)] [string] $SourceEwsFqdn,         # e.g. mail.source.tld
  [Parameter(Mandatory=$true)] [string] $SourceDC,              # e.g. dc01.source.local
  [Parameter(Mandatory=$true)] [string] $SourceAuthoritativeDomain, # e.g. source.tld
  
  # === Target Environment Configuration ===
  [Parameter(Mandatory=$true)] [string] $TargetDC,              # e.g. dc01.target.local
  [Parameter(Mandatory=$true)] [string] $TargetDeliveryDomain,  # e.g. target.tld
  
  # === User Lists ===
  [Parameter(Mandatory=$true)] [string] $UsersPrepareCsvPath,   # users_prepare.csv (Column: Identity)
  
  # === Optional Free/Busy Configuration ===
  [Parameter()] [switch] $ConfigureFreeBusy,
  [Parameter()] [string] $OrgRelationshipName = "SourceOrg",
  [Parameter()] [string] $SourceFederatedDomain,
  
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
    UsersProcessed = 0
    UsersSuccessful = 0
    UsersFailed = 0
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
    $saveDialog.FileName = "Target-Prepare-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
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
    <title>Exchange Migration - Target Environment Preparation Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f7fa; color: #333; line-height: 1.6; }
        .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
        .header { background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%); color: white; padding: 30px; border-radius: 10px; margin-bottom: 30px; text-align: center; }
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
        .progress-fill { height: 100%; background: linear-gradient(90deg, #3498db, #2980b9); transition: width 0.3s ease; }
        .user-stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; margin: 20px 0; }
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
            <p>Target Environment Preparation - Step 2</p>
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
                <div class="stat-number success">$($script:Statistics.UsersSuccessful)</div>
                <div class="stat-label">Users Prepared</div>
            </div>
            <div class="stat-card">
                <div class="stat-number error">$($script:Statistics.UsersFailed)</div>
                <div class="stat-label">Failed Users</div>
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
                    <option value="Validation">Validation</option>
                    <option value="User Preparation">User Preparation</option>
                    <option value="Free/Busy Setup">Free/Busy Setup</option>
                    <option value="Connectivity Test">Connectivity Test</option>
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
    Write-Header "Target Environment Preparation - Step 2"
    Write-LogEntry -Message "Starting Exchange Cross-Forest Migration - Target Environment Preparation" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Source EWS FQDN: $SourceEwsFqdn" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Target DC: $TargetDC" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Target Delivery Domain: $TargetDeliveryDomain" -Level "Info" -Phase "General"
    
    # Create output directory
    if (-not (Test-Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-LogEntry -Message "Created output directory: $OutputPath" -Level "Info" -Phase "General"
    }

    # Get credentials
    Write-Header "Credential Collection"
    $LocalCredentials  = Get-Credential -Message "Target Forest (LOCAL) Admin"
    $RemoteCredentials = Get-Credential -Message "Source Forest (REMOTE) Admin"
    Write-LogEntry -Message "Credentials collected successfully" -Level "Success" -Phase "General"
}

process {
    try {
        # Step 1: Validate CSV file
        Write-Header "CSV Validation"
        Write-LogEntry -Message "Validating user preparation CSV file: $UsersPrepareCsvPath" -Level "Info" -Phase "Validation"
        
        if (-not (Test-Path $UsersPrepareCsvPath)) {
            throw "CSV file not found: $UsersPrepareCsvPath"
        }
        
        $users = Import-Csv $UsersPrepareCsvPath
        if (-not $users -or $users.Count -eq 0) {
            throw "CSV file $UsersPrepareCsvPath contains no entries"
        }
        
        if (-not ($users | Get-Member -Name "Identity")) {
            throw "CSV file must contain 'Identity' column"
        }
        
        $validUsers = $users | Where-Object { $_.Identity -and $_.Identity.Trim() -ne "" }
        if ($validUsers.Count -eq 0) {
            throw "No valid Identity entries found in CSV file"
        }
        
        Write-LogEntry -Message "Found $($validUsers.Count) valid users to prepare" -Level "Success" -Phase "Validation"
        $script:Statistics.CompletedSteps++

        # Step 2: Validate Prepare-MoveRequest.ps1 script
        Write-Header "Script Validation"
        Write-LogEntry -Message "Validating Prepare-MoveRequest.ps1 script availability" -Level "Info" -Phase "Validation"
        
        $scriptPath = Join-Path $env:ExchangeInstallPath "Scripts\Prepare-MoveRequest.ps1"
        if (-not (Test-Path $scriptPath)) {
            throw "Prepare-MoveRequest.ps1 not found at $scriptPath"
        }
        
        Write-LogEntry -Message "Prepare-MoveRequest.ps1 script found at: $scriptPath" -Level "Success" -Phase "Validation"
        $script:Statistics.CompletedSteps++

        # Step 3: Test connectivity to source
        Write-Header "Source Connectivity Test"
        Write-LogEntry -Message "Testing connectivity to source environment" -Level "Info" -Phase "Connectivity Test"
        
        try {
            Test-MigrationServerAvailability -ExchangeRemoteMove -RemoteServer $SourceEwsFqdn -Credentials $RemoteCredentials -ErrorAction Stop | Out-Null
            Write-LogEntry -Message "Source connectivity test successful" -Level "Success" -Phase "Connectivity Test"
        } catch {
            Write-LogEntry -Message "Source connectivity test failed: $($_.Exception.Message)" -Level "Warning" -Phase "Connectivity Test"
            Write-LogEntry -Message "Continuing with user preparation - connectivity will be tested during migration" -Level "Info" -Phase "Connectivity Test"
        }
        $script:Statistics.CompletedSteps++

        # Step 4: Prepare MailUser objects
        Write-Header "MailUser Preparation"
        Write-LogEntry -Message "Starting MailUser preparation for $($validUsers.Count) users" -Level "Info" -Phase "User Preparation"
        
        foreach ($user in $validUsers) {
            $identity = $user.Identity.Trim()
            $script:Statistics.UsersProcessed++
            
            Write-LogEntry -Message "Preparing MailUser for: $identity" -Level "Info" -Phase "User Preparation"
            
            try {
                & $scriptPath -Identity $identity `
                    -RemoteForestDomainController $SourceDC `
                    -RemoteForestCredential $RemoteCredentials `
                    -LocalForestDomainController $TargetDC `
                    -LocalForestCredential $LocalCredentials `
                    -MailboxDeliveryDomain $SourceAuthoritativeDomain `
                    -ErrorAction Stop
                
                Write-LogEntry -Message "Successfully prepared MailUser for: $identity" -Level "Success" -Phase "User Preparation"
                $script:Statistics.UsersSuccessful++
            } catch {
                Write-LogEntry -Message "Failed to prepare MailUser for $identity`: $($_.Exception.Message)" -Level "Error" -Phase "User Preparation"
                $script:Statistics.UsersFailed++
            }
        }
        
        Write-LogEntry -Message "MailUser preparation completed. Success: $($script:Statistics.UsersSuccessful), Failed: $($script:Statistics.UsersFailed)" -Level "Info" -Phase "User Preparation"
        $script:Statistics.CompletedSteps++

        # Step 5: Optional Free/Busy Configuration
        if ($ConfigureFreeBusy) {
            Write-Header "Free/Busy Configuration"
            Write-LogEntry -Message "Configuring Free/Busy coexistence" -Level "Info" -Phase "Free/Busy Setup"
            
            try {
                $federatedDomain = if ($SourceFederatedDomain) { $SourceFederatedDomain } else { $SourceAuthoritativeDomain }
                
                Get-FederationInformation -DomainName $federatedDomain | 
                    New-OrganizationRelationship -Name $OrgRelationshipName `
                        -FreeBusyAccessEnabled $true -FreeBusyAccessLevel LimitedDetails -ErrorAction Stop | Out-Null
                
                Write-LogEntry -Message "Organization Relationship '$OrgRelationshipName' created successfully" -Level "Success" -Phase "Free/Busy Setup"
            } catch {
                Write-LogEntry -Message "Failed to create Organization Relationship: $($_.Exception.Message)" -Level "Warning" -Phase "Free/Busy Setup"
                Write-LogEntry -Message "Free/Busy setup requires proper federation configuration" -Level "Info" -Phase "Free/Busy Setup"
            }
        } else {
            Write-LogEntry -Message "Free/Busy configuration skipped (not requested)" -Level "Info" -Phase "Free/Busy Setup"
        }
        $script:Statistics.CompletedSteps++

        # Display next steps
        Write-Header "Next Steps"
        Write-LogEntry -Message "Target environment preparation completed successfully" -Level "Success" -Phase "General"
        Write-LogEntry -Message "Next: Run script 03-Target-Environment-Migration.ps1 on the TARGET environment" -Level "Info" -Phase "General"
        Write-LogEntry -Message "Ensure users_migrate.csv is prepared with EmailAddress column" -Level "Info" -Phase "General"
        
        $script:Statistics.Success = $true

    } catch {
        Write-LogEntry -Message "Target environment preparation failed: $($_.Exception.Message)" -Level "Error" -Phase "General"
        $script:Statistics.Success = $false
        throw
    }
}

end {
    $script:Statistics.EndTime = Get-Date
    $duration = $script:Statistics.EndTime - $script:Statistics.StartTime
    
    Write-Header "Target Environment Preparation Complete"
    Write-LogEntry -Message "Total execution time: $($duration.ToString('hh\:mm\:ss'))" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Steps completed: $($script:Statistics.CompletedSteps)/$($script:Statistics.TotalSteps)" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Users processed: $($script:Statistics.UsersProcessed), Successful: $($script:Statistics.UsersSuccessful), Failed: $($script:Statistics.UsersFailed)" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Errors: $($script:Statistics.Errors), Warnings: $($script:Statistics.Warnings)" -Level "Info" -Phase "General"
    
    # Generate HTML report
    $reportHtml = Generate-HTMLReport -OutputPath $OutputPath
    
    # Save report
    $reportPath = $null
    if ($ShowSaveDialog) {
        $reportPath = Show-SaveDialog -DefaultPath $OutputPath
    }
    
    if (-not $reportPath) {
        $reportPath = Join-Path $OutputPath "Target-Prepare-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    }
    
    try {
        $reportHtml | Out-File -FilePath $reportPath -Encoding UTF8
        Write-LogEntry -Message "HTML report saved to: $reportPath" -Level "Success" -Phase "General"
        
        # Open report in default browser
        Start-Process $reportPath
    } catch {
        Write-LogEntry -Message "Failed to save HTML report: $($_.Exception.Message)" -Level "Error" -Phase "General"
    }
    
    Write-Host "`nTarget environment preparation completed!" -ForegroundColor Green
    Write-Host "Report saved to: $reportPath" -ForegroundColor Yellow
    Write-Host "`nNext step: Run 03-Target-Environment-Migration.ps1 on the TARGET environment" -ForegroundColor Cyan
}
