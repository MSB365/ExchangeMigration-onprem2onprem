<#
.SYNOPSIS
  Cross-Forest Exchange Migration - TARGET Environment Migration (Step 3)
.DESCRIPTION
  Creates migration endpoints, batches, and manages the actual mailbox migration process.
  Run this script on the TARGET Exchange environment after Step 2.
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
  [Parameter(Mandatory=$true)] [string[]] $TargetDatabases,     # Target mailbox databases
  
  # === Migration Configuration ===
  [Parameter(Mandatory=$true)] [string] $UsersMigrateCsvPath,   # users_migrate.csv (Column: EmailAddress)
  [Parameter()] [string] $EndpointName = "CF-Source-EWS",
  [Parameter()] [string] $BatchPrefix = "CF-Batch",
  [Parameter()] [int] $BatchSize = 100,
  [Parameter()] [int] $BadItemLimit = 10,
  [Parameter()] [int] $LargeItemLimit = 10,
  [Parameter()] [switch] $SuspendWhenReadyToComplete,
  [Parameter()] [datetime] $CompleteAtLocalTime,
  [Parameter()] [ValidateSet("Pilot", "Bulk")] [string] $MigrationType = "Bulk",
  
  # === Reporting ===
  [Parameter()] [string] $OutputPath = ".\MigrationReports",
  [Parameter()] [switch] $ShowSaveDialog = $true
)

# Initialize logging and statistics
$script:LogEntries = @()
$script:Statistics = @{
    StartTime = Get-Date
    EndTime = $null
    TotalSteps = 6
    CompletedSteps = 0
    BatchesCreated = 0
    UsersInMigration = 0
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
    $saveDialog.FileName = "Target-Migration-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    $saveDialog.InitialDirectory = $DefaultPath
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $saveDialog.FileName
    }
    return $null
}

function Get-NextDatabase {
    param([string[]] $Databases)
    
    if (-not $script:DatabaseEnumerator) {
        $script:DatabaseEnumerator = [System.Collections.Generic.Queue[string]]::new()
        $Databases | ForEach-Object { $script:DatabaseEnumerator.Enqueue($_) }
    }
    
    $db = $script:DatabaseEnumerator.Dequeue()
    $script:DatabaseEnumerator.Enqueue($db)
    return $db
}

function Generate-HTMLReport {
    param([string] $OutputPath)
    
    $reportHtml = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exchange Migration - Target Environment Migration Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f7fa; color: #333; line-height: 1.6; }
        .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
        .header { background: linear-gradient(135deg, #27ae60 0%, #2ecc71 100%); color: white; padding: 30px; border-radius: 10px; margin-bottom: 30px; text-align: center; }
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
            <p>Target Environment Migration - Step 3</p>
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
                <div class="stat-number success">$($script:Statistics.BatchesCreated)</div>
                <div class="stat-label">Batches Created</div>
            </div>
            <div class="stat-card">
                <div class="stat-number info">$($script:Statistics.UsersInMigration)</div>
                <div class="stat-label">Users in Migration</div>
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
                    <option value="Endpoint Creation">Endpoint Creation</option>
                    <option value="Batch Creation">Batch Creation</option>
                    <option value="Migration Monitoring">Migration Monitoring</option>
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
    Write-Header "Target Environment Migration - Step 3"
    Write-LogEntry -Message "Starting Exchange Cross-Forest Migration - Target Environment Migration" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Migration Type: $MigrationType" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Source EWS FQDN: $SourceEwsFqdn" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Target Databases: $($TargetDatabases -join ', ')" -Level "Info" -Phase "General"
    
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

    # Calculate CompleteAfter time if specified
    if (-not $SuspendWhenReadyToComplete -and $CompleteAtLocalTime) {
        $CompleteAfterUtc = $CompleteAtLocalTime.ToUniversalTime()
        Write-LogEntry -Message "Migration completion scheduled for: $CompleteAfterUtc UTC" -Level "Info" -Phase "General"
    }
}

process {
    try {
        # Step 1: Validate CSV file
        Write-Header "CSV Validation"
        Write-LogEntry -Message "Validating migration CSV file: $UsersMigrateCsvPath" -Level "Info" -Phase "Validation"
        
        if (-not (Test-Path $UsersMigrateCsvPath)) {
            throw "CSV file not found: $UsersMigrateCsvPath"
        }
        
        $users = Import-Csv $UsersMigrateCsvPath
        if (-not $users -or $users.Count -eq 0) {
            throw "CSV file $UsersMigrateCsvPath contains no entries"
        }
        
        if (-not ($users | Get-Member -Name "EmailAddress")) {
            throw "CSV file must contain 'EmailAddress' column"
        }
        
        $validUsers = $users | Where-Object { $_.EmailAddress -and $_.EmailAddress.Trim() -ne "" }
        if ($validUsers.Count -eq 0) {
            throw "No valid EmailAddress entries found in CSV file"
        }
        
        # Filter for pilot if specified
        if ($MigrationType -eq "Pilot") {
            $validUsers = $validUsers | Select-Object -First $BatchSize
            Write-LogEntry -Message "Pilot migration: Limited to first $($validUsers.Count) users" -Level "Info" -Phase "Validation"
        }
        
        Write-LogEntry -Message "Found $($validUsers.Count) valid users for migration" -Level "Success" -Phase "Validation"
        $script:Statistics.UsersInMigration = $validUsers.Count
        $script:Statistics.CompletedSteps++

        # Step 2: Test source connectivity
        Write-Header "Source Connectivity Test"
        Write-LogEntry -Message "Testing connectivity to source environment" -Level "Info" -Phase "Connectivity Test"
        
        try {
            Test-MigrationServerAvailability -ExchangeRemoteMove -RemoteServer $SourceEwsFqdn -Credentials $RemoteCredentials -ErrorAction Stop | Out-Null
            Write-LogEntry -Message "Source connectivity test successful" -Level "Success" -Phase "Connectivity Test"
        } catch {
            Write-LogEntry -Message "Source connectivity test failed: $($_.Exception.Message)" -Level "Error" -Phase "Connectivity Test"
            throw "Cannot proceed without source connectivity"
        }
        $script:Statistics.CompletedSteps++

        # Step 3: Create or validate migration endpoint
        Write-Header "Migration Endpoint"
        Write-LogEntry -Message "Creating/validating migration endpoint: $EndpointName" -Level "Info" -Phase "Endpoint Creation"
        
        $endpoint = Get-MigrationEndpoint -Identity $EndpointName -ErrorAction SilentlyContinue
        if (-not $endpoint) {
            try {
                $endpoint = New-MigrationEndpoint -Name $EndpointName -ExchangeRemoteMove `
                    -RemoteServer $SourceEwsFqdn -Credentials $RemoteCredentials -ErrorAction Stop
                Write-LogEntry -Message "Migration endpoint '$EndpointName' created successfully" -Level "Success" -Phase "Endpoint Creation"
            } catch {
                Write-LogEntry -Message "Failed to create migration endpoint: $($_.Exception.Message)" -Level "Error" -Phase "Endpoint Creation"
                throw
            }
        } else {
            Write-LogEntry -Message "Migration endpoint '$EndpointName' already exists" -Level "Info" -Phase "Endpoint Creation"
        }
        $script:Statistics.CompletedSteps++

        # Step 4: Create migration batches
        Write-Header "Migration Batch Creation"
        Write-LogEntry -Message "Creating migration batches with batch size: $BatchSize" -Level "Info" -Phase "Batch Creation"
        
        # Split users into batches
        $userBatches = @()
        for ($i = 0; $i -lt $validUsers.Count; $i += $BatchSize) {
            $endIndex = [Math]::Min($i + $BatchSize - 1, $validUsers.Count - 1)
            $userBatches += ,($validUsers[$i..$endIndex])
        }
        
        Write-LogEntry -Message "Will create $($userBatches.Count) migration batches" -Level "Info" -Phase "Batch Creation"
        
        $batchIndex = 1
        foreach ($batch in $userBatches) {
            $batchName = "{0}-{1:000}" -f $BatchPrefix, $batchIndex
            $batchIndex++
            
            Write-LogEntry -Message "Creating batch '$batchName' with $($batch.Count) users" -Level "Info" -Phase "Batch Creation"
            
            try {
                # Create temporary CSV for this batch
                $tempCsv = New-TemporaryFile
                "EmailAddress" | Out-File -FilePath $tempCsv -Encoding UTF8
                $batch | ForEach-Object { $_.EmailAddress } | Out-File -FilePath $tempCsv -Append -Encoding UTF8
                $csvBytes = [System.IO.File]::ReadAllBytes($tempCsv)
                
                # Get target database using round-robin
                $targetDb = Get-NextDatabase -Databases $TargetDatabases
                
                # Create batch parameters
                $batchParams = @{
                    Name = $batchName
                    CSVData = $csvBytes
                    SourceEndpoint = $EndpointName
                    TargetDatabases = $targetDb
                    AutoStart = $true
                    BadItemLimit = $BadItemLimit
                    LargeItemLimit = $LargeItemLimit
                }
                
                # Add completion settings
                if ($SuspendWhenReadyToComplete) {
                    $batchParams.SuspendWhenReadyToComplete = $true
                } elseif ($CompleteAfterUtc) {
                    $batchParams.CompleteAfter = $CompleteAfterUtc
                }
                
                # Create the batch
                $newBatch = New-MigrationBatch @batchParams
                Write-LogEntry -Message "Batch '$batchName' created successfully -> Target DB: $targetDb" -Level "Success" -Phase "Batch Creation"
                $script:Statistics.BatchesCreated++
                
                # Clean up temp file
                Remove-Item $tempCsv -Force -ErrorAction SilentlyContinue
                
            } catch {
                Write-LogEntry -Message "Failed to create batch '$batchName': $($_.Exception.Message)" -Level "Error" -Phase "Batch Creation"
                # Continue with other batches
            }
        }
        $script:Statistics.CompletedSteps++

        # Step 5: Initial migration monitoring
        Write-Header "Migration Monitoring"
        Write-LogEntry -Message "Displaying initial migration status" -Level "Info" -Phase "Migration Monitoring"
        
        try {
            $batches = Get-MigrationBatch | Where-Object { $_.Name -like "$BatchPrefix*" }
            foreach ($batch in $batches) {
                Write-LogEntry -Message "Batch: $($batch.Name), Status: $($batch.Status), Users: $($batch.TotalCount), Progress: $($batch.PercentageComplete)%" -Level "Info" -Phase "Migration Monitoring"
            }
            
            $moveRequests = Get-MoveRequest | Get-MoveRequestStatistics | Select-Object -First 10
            foreach ($mr in $moveRequests) {
                Write-LogEntry -Message "Move Request: $($mr.DisplayName), Status: $($mr.StatusDetail), Progress: $($mr.PercentComplete)%" -Level "Info" -Phase "Migration Monitoring"
            }
        } catch {
            Write-LogEntry -Message "Error retrieving migration status: $($_.Exception.Message)" -Level "Warning" -Phase "Migration Monitoring"
        }
        $script:Statistics.CompletedSteps++

        # Step 6: Display next steps
        Write-Header "Next Steps"
        Write-LogEntry -Message "Migration batches created and started successfully" -Level "Success" -Phase "General"
        Write-LogEntry -Message "Monitor progress using Get-MigrationBatch and Get-MoveRequestStatistics" -Level "Info" -Phase "General"
        
        if ($SuspendWhenReadyToComplete) {
            Write-LogEntry -Message "Batches will suspend at 95% completion. Use 04-Target-Environment-Finalize.ps1 to complete" -Level "Info" -Phase "General"
        } else {
            Write-LogEntry -Message "Batches will complete automatically. Use 04-Target-Environment-Finalize.ps1 for cleanup" -Level "Info" -Phase "General"
        }
        
        $script:Statistics.CompletedSteps++
        $script:Statistics.Success = $true

    } catch {
        Write-LogEntry -Message "Migration setup failed: $($_.Exception.Message)" -Level "Error" -Phase "General"
        $script:Statistics.Success = $false
        throw
    }
}

end {
    $script:Statistics.EndTime = Get-Date
    $duration = $script:Statistics.EndTime - $script:Statistics.StartTime
    
    Write-Header "Target Environment Migration Setup Complete"
    Write-LogEntry -Message "Total execution time: $($duration.ToString('hh\:mm\:ss'))" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Steps completed: $($script:Statistics.CompletedSteps)/$($script:Statistics.TotalSteps)" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Batches created: $($script:Statistics.BatchesCreated), Users in migration: $($script:Statistics.UsersInMigration)" -Level "Info" -Phase "General"
    Write-LogEntry -Message "Errors: $($script:Statistics.Errors), Warnings: $($script:Statistics.Warnings)" -Level "Info" -Phase "General"
    
    # Generate HTML report
    $reportHtml = Generate-HTMLReport -OutputPath $OutputPath
    
    # Save report
    $reportPath = $null
    if ($ShowSaveDialog) {
        $reportPath = Show-SaveDialog -DefaultPath $OutputPath
    }
    
    if (-not $reportPath) {
        $reportPath = Join-Path $OutputPath "Target-Migration-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    }
    
    try {
        $reportHtml | Out-File -FilePath $reportPath -Encoding UTF8
        Write-LogEntry -Message "HTML report saved to: $reportPath" -Level "Success" -Phase "General"
        
        # Open report in default browser
        Start-Process $reportPath
    } catch {
        Write-LogEntry -Message "Failed to save HTML report: $($_.Exception.Message)" -Level "Error" -Phase "General"
    }
    
    Write-Host "`nMigration setup completed!" -ForegroundColor Green
    Write-Host "Report saved to: $reportPath" -ForegroundColor Yellow
    Write-Host "`nNext step: Monitor migration progress and run 04-Target-Environment-Finalize.ps1 when ready" -ForegroundColor Cyan
}
