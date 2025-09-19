Thought for 1sRead complete READMEHere's the complete README.md file that was created:

# Exchange Cross-Forest Migration Toolkit

A comprehensive PowerShell toolkit for migrating Exchange mailboxes between Active Directory forests with trust relationships. This toolkit provides a structured, phase-based approach to cross-forest migrations with detailed HTML reporting and error handling.

## Overview

This toolkit splits the complex migration process into **6 numbered scripts** that must be executed in a specific order across **Source** and **Target** environments. Each script includes comprehensive HTML reporting with save dialog functionality.



## Script Execution Order

### Phase 1: Environment Setup
| Step | Script | Environment | Description |
|------|--------|-------------|-------------|
| **1** | `01-Source-Environment-Setup.ps1` | **SOURCE** | Enable MRS-Proxy, test connectivity |
| **2** | `02-Target-Environment-Prepare.ps1` | **TARGET** | Prepare MailUser objects |

### Phase 2: Migration Execution  
| Step | Script | Environment | Description |
|------|--------|-------------|-------------|
| **3** | `03-Target-Environment-Migration.ps1` | **TARGET** | Create endpoints, batches, start migration |
| **4** | `04-Target-Environment-Finalize.ps1` | **TARGET** | Finalize batches, resume suspended requests |

### Phase 3: Cleanup
| Step | Script | Environment | Description |
|------|--------|-------------|-------------|
| **5** | `05-Target-Environment-Cleanup.ps1` | **TARGET** | Remove completed batches, move requests |
| **6** | `06-Source-Environment-Teardown.ps1` | **SOURCE** | Disable MRS-Proxy, final cleanup |

## Quick Start

### Prerequisites

### Active Directory Trust
- **Forest Trust** or **External Trust** established between source and target forests
- **DNS resolution** working bidirectionally
- **Time synchronization** between forests

### Exchange Requirements
- **Exchange 2013** or later in both environments
- **Organization Management** rights in both forests
- **MRS-Proxy** capability on source Exchange servers
- **Mailbox databases** created in target environment

### Network Requirements
- **HTTPS connectivity** from target to source (port 443)
- **EWS endpoint** accessible: `https://source-server/EWS/mrsproxy.svc`
- **Firewall rules** allowing Exchange Web Services traffic

### PowerShell Requirements
- **Exchange Management Shell** on both environments
- **PowerShell 5.1** or later
- **System.Windows.Forms** assembly (for save dialogs)

## CSV File Formats

### users_prepare.csv (Step 2)

```csv
Identity
john.doe@source.tld
jane.smith@source.tld
admin@source.tld
```

### users_migrate.csv (Step 3)
```csv
EmailAddress
john.doe@source.tld
jane.smith@source.tld
admin@source.tld
```


### Migration Overview

This toolkit splits the migration process into **6 numbered scripts** that run on specific Exchange environments:

| Script | Environment | Purpose | Duration
|-----|-----|-----|-----
| **01** | SOURCE | Setup MRS-Proxy | 5-10 minutes
| **02** | TARGET | Prepare MailUsers | 30-60 minutes
| **03** | TARGET | Start Migration | 1-2 hours
| **04** | TARGET | Finalize Migration | 30-60 minutes
| **05** | TARGET | Cleanup Artifacts | 10-20 minutes
| **06** | SOURCE | Teardown Setup | 5-10 minutes


## Required CSV Files

### users_prepare.csv

```csv

  0
  Identityjohn.doe@source.tldjane.smith@source.tldmike.johnson@source.tld


```

### users_migrate.csv

```csv

  0
  EmailAddressjohn.doe@source.tldjane.smith@source.tldmike.johnson@source.tld


```

## Step-by-Step Execution Guide

### Step 1: Source Environment Setup

**Environment:** SOURCE Exchange Management Shell**Script:** `01-Source-Environment-Setup.ps1`

```powershell
.\01-Source-Environment-Setup.ps1 `
    -SourceEwsFqdn "mail.source.tld" `
    -RestartIIS
```

**What it does:**

- Enables MRS-Proxy on all EWS virtual directories
- Restarts IIS to apply changes
- Tests MRS-Proxy endpoint accessibility
- Validates configuration


**HTML Report:** Source environment setup status and connectivity tests

---

### Step 2: Target Environment Preparation

**Environment:** TARGET Exchange Management Shell**Script:** `02-Target-Environment-Prepare.ps1`

```powershell
.\02-Target-Environment-Prepare.ps1 `
    -SourceEwsFqdn "mail.source.tld" `
    -SourceDC "dc01.source.local" `
    -SourceAuthoritativeDomain "source.tld" `
    -TargetDC "dc01.target.local" `
    -TargetDeliveryDomain "target.tld" `
    -UsersPrepareCsvPath ".\users_prepare.csv" `
    -ConfigureFreeBusy
```

**What it does:**

- Validates CSV file format and content
- Runs `Prepare-MoveRequest.ps1` for each user
- Creates MailUser objects in target forest
- Optionally configures Free/Busy coexistence


**HTML Report:** User preparation results, success/failure counts

---

### Step 3: Target Environment Migration

**Environment:** TARGET Exchange Management Shell**Script:** `03-Target-Environment-Migration.ps1`

```powershell
# Pilot Migration (recommended first)
.\03-Target-Environment-Migration.ps1 `
    -SourceEwsFqdn "mail.source.tld" `
    -SourceDC "dc01.source.local" `
    -SourceAuthoritativeDomain "source.tld" `
    -TargetDC "dc01.target.local" `
    -TargetDeliveryDomain "target.tld" `
    -TargetDatabases @("DB01", "DB02", "DB03") `
    -UsersMigrateCsvPath ".\users_migrate.csv" `
    -MigrationType "Pilot" `
    -BatchSize 10 `
    -SuspendWhenReadyToComplete

# Full Migration
.\03-Target-Environment-Migration.ps1 `
    -SourceEwsFqdn "mail.source.tld" `
    -SourceDC "dc01.source.local" `
    -SourceAuthoritativeDomain "source.tld" `
    -TargetDC "dc01.target.local" `
    -TargetDeliveryDomain "target.tld" `
    -TargetDatabases @("DB01", "DB02", "DB03") `
    -UsersMigrateCsvPath ".\users_migrate.csv" `
    -MigrationType "Bulk" `
    -BatchSize 100 `
    -CompleteAtLocalTime "2024-01-15 02:00:00"
```

**What it does:**

- Creates migration endpoint to source
- Splits users into batches with round-robin database assignment
- Creates and starts migration batches
- Provides initial migration status


**HTML Report:** Batch creation status, migration progress, database assignments

---

### Step 4: Target Environment Finalization

**Environment:** TARGET Exchange Management Shell**Script:** `04-Target-Environment-Finalize.ps1`

```powershell
# Finalize synced batches
.\04-Target-Environment-Finalize.ps1 -Action "Finalize"

# Resume suspended move requests
.\04-Target-Environment-Finalize.ps1 -Action "Resume"

# Monitor migration progress
.\04-Target-Environment-Finalize.ps1 -Action "Monitor"
```

**What it does:**

- Finalizes migration batches in "Synced" status
- Resumes suspended move requests
- Monitors migration progress and status
- Provides detailed completion statistics


**HTML Report:** Finalization results, completion status, migration statistics

---

### Step 5: Target Environment Cleanup

**Environment:** TARGET Exchange Management Shell**Script:** `05-Target-Environment-Cleanup.ps1`

```powershell
.\05-Target-Environment-Cleanup.ps1 `
    -BatchPrefix "CF-Batch" `
    -RemoveEndpoint `
    -RemoveOrgRelationship `
    -EndpointName "CF-Source-EWS" `
    -OrgRelationshipName "SourceOrg"
```

**What it does:**

- Removes completed move requests
- Removes completed migration batches
- Optionally removes migration endpoints
- Optionally removes organization relationships
- Verifies cleanup completion


**HTML Report:** Cleanup results, remaining artifacts, verification status

---

### Step 6: Source Environment Teardown

**Environment:** SOURCE Exchange Management Shell**Script:** `06-Source-Environment-Teardown.ps1`

```powershell
.\06-Source-Environment-Teardown.ps1 `
    -SourceEwsFqdn "mail.source.tld" `
    -DisableMRSProxy `
    -RestartIIS
```

**What it does:**

- Disables MRS-Proxy on all EWS virtual directories
- Restarts IIS to apply changes
- Verifies MRS-Proxy is no longer accessible
- Completes migration project teardown


**HTML Report:** Teardown results, final verification, project completion

## HTML Reporting Features

Each script generates comprehensive HTML reports with:

### Interactive Features

- **Search functionality** - Filter logs by keyword
- **Phase filtering** - View logs by execution phase
- **Level filtering** - Filter by Info, Warning, Error, Success
- **Responsive design** - Works on desktop and mobile


### Visual Elements

- **Progress bars** - Visual execution progress
- **Statistics cards** - Key metrics and counts
- **Color-coded logs** - Easy identification of issues
- **Professional styling** - Clean, modern interface


### Save Dialog

- **Popup save dialog** - Choose custom save location
- **Auto-naming** - Timestamped report names
- **Auto-open** - Reports open in default browser
- **Default location** - Falls back to `.\MigrationReports`


## Migration Monitoring

### PowerShell Commands

```powershell
# Check migration batch status
Get-MigrationBatch | Where-Object {$_.Name -like "CF-Batch*"} | 
    Format-Table Name, Status, TotalCount, PercentageComplete

# Check move request status
Get-MoveRequest | Get-MoveRequestStatistics | 
    Format-Table DisplayName, Status, PercentComplete, BytesTransferred

# Check for errors
Get-MoveRequest | Where-Object {$_.Status -eq "Failed"} | 
    Get-MoveRequestStatistics | Format-List
```

### Migration States

| Status | Description | Action Required
|-----|-----|-----|-----
| **Queued** | Waiting to start | Monitor progress
| **InProgress** | Actively migrating | Monitor progress
| **Synced** | 95% complete, ready to finalize | Run Step 4 (Finalize)
| **AutoSuspended** | Suspended at 95% | Run Step 4 (Resume)
| **Completed** | Successfully finished | Run Step 5 (Cleanup)
| **Failed** | Migration failed | Investigate and retry


## Advanced Configuration

### Batch Size Optimization

```powershell
# Small batches for testing
-BatchSize 10

# Medium batches for production
-BatchSize 50

# Large batches for bulk migration
-BatchSize 100
```

### Migration Timing Control

```powershell
# Suspend at 95% for manual completion
-SuspendWhenReadyToComplete

# Automatic completion at specific time
-CompleteAtLocalTime "2024-01-15 02:00:00"
```

### Database Distribution

```powershell
# Single database
-TargetDatabases @("DB01")

# Multiple databases for load balancing
-TargetDatabases @("DB01", "DB02", "DB03", "DB04")
```

## Troubleshooting

### Common Issues

#### MRS-Proxy Connection Failed

```powershell
# Verify MRS-Proxy is enabled
Get-WebServicesVirtualDirectory | Format-List Name, MRSProxyEnabled

# Test connectivity manually
Test-MigrationServerAvailability -ExchangeRemoteMove `
    -RemoteServer "mail.source.tld" -Credentials $cred
```

#### Prepare-MoveRequest.ps1 Not Found

```powershell
# Check Exchange installation path
$env:ExchangeInstallPath
# Script should be at: $env:ExchangeInstallPath\Scripts\Prepare-MoveRequest.ps1
```

#### Permission Denied Errors

- Ensure running account has **Organization Management** rights
- Verify **cross-forest trust** is working properly
- Check **DNS resolution** between forests
- Validate **time synchronization** between domains


#### Migration Stuck or Slow

```powershell
# Check for large items or corruption
Get-MoveRequestStatistics -Identity "user@domain.com" -IncludeReport | 
    Select-Object -ExpandProperty Report

# Resume suspended requests
Get-MoveRequest | Where-Object {$_.Status -eq "Suspended"} | Resume-MoveRequest
```

### Log Analysis

- Use HTML report **search functionality** to find specific errors
- Filter by **Error** level to isolate issues quickly
- Check **Exchange event logs** on both source and target servers
- Review **IIS logs** for EWS/MRS-Proxy connectivity issues


## Security Considerations

### Credential Management

- Scripts prompt for credentials **interactively** (no storage)
- Use **dedicated migration accounts** with minimal required permissions
- Consider **service accounts** for unattended execution


### Network Security

- Ensure **encrypted communication** (HTTPS/TLS) between forests
- Implement **firewall rules** for Exchange Web Services
- Monitor **network traffic** during migration for anomalies


### Audit Trail

- All actions are **logged comprehensively** for compliance
- HTML reports provide **complete audit trail**
- Consider **archiving reports** for long-term retention


### Cleanup Security

- **Remove migration artifacts** after completion (Step 5)
- **Disable MRS-Proxy** when no longer needed (Step 6)
- **Review permissions** granted during migration process


## Performance Optimization

### Batch Sizing Guidelines

- **Pilot migrations**: 5-10 users for testing
- **Production migrations**: 50-100 users per batch
- **Large environments**: Consider multiple parallel batches


### Database Distribution

- Use **multiple target databases** for load distribution
- **Round-robin assignment** spreads load evenly
- Monitor **database performance** during migration


### Timing Considerations

- Schedule migrations during **off-peak hours**
- Use **CompleteAtLocalTime** for controlled cutover
- Consider **SuspendWhenReadyToComplete** for manual control


## Best Practices

### Pre-Migration

1. **Test in lab environment** before production
2. **Verify all prerequisites** are met
3. **Create proper backups** of both environments
4. **Document current configuration** for rollback


### During Migration

1. **Start with pilot group** (Step 3 with MigrationType "Pilot")
2. **Monitor progress regularly** using HTML reports
3. **Address issues promptly** before proceeding
4. **Communicate status** to stakeholders


### Post-Migration

1. **Verify user access** and functionality
2. **Test mail flow** in both directions
3. **Clean up artifacts** promptly (Steps 5-6)
4. **Archive reports** for compliance


## Parameter Reference

### Common Parameters

| Parameter | Description | Example
|-----|-----|-----|-----
| `SourceEwsFqdn` | Source Exchange Web Services FQDN | `mail.source.tld`
| `SourceDC` | Source domain controller | `dc01.source.local`
| `SourceAuthoritativeDomain` | Source authoritative domain | `source.tld`
| `TargetDC` | Target domain controller | `dc01.target.local`
| `TargetDeliveryDomain` | Target delivery domain | `target.tld`
| `TargetDatabases` | Array of target databases | `@("DB01","DB02")`
| `BatchSize` | Users per migration batch | `100`
| `ConfigureFreeBusy` | Enable Free/Busy coexistence | Switch parameter


### Migration Control Parameters

| Parameter | Description | Example
|-----|-----|-----|-----
| `MigrationType` | Pilot or Bulk migration | `"Pilot"` or `"Bulk"`
| `SuspendWhenReadyToComplete` | Suspend at 95% | Switch parameter
| `CompleteAtLocalTime` | Scheduled completion | `"2024-01-15 02:00:00"`
| `BadItemLimit` | Bad items per mailbox | `10`
| `LargeItemLimit` | Large items per mailbox | `10`


## Support and Troubleshooting

### Script Execution Issues

- Ensure **PowerShell execution policy** allows script execution
- Run **Exchange Management Shell** as administrator
- Verify **Exchange cmdlets** are available


### Migration Issues

- Check **Exchange event logs** for detailed error information
- Use **Get-MoveRequestStatistics** for detailed migration status
- Review **IIS logs** for EWS connectivity issues


### Reporting Issues

- Ensure **System.Windows.Forms** assembly is available
- Check **file permissions** in output directory
- Verify **default browser** is configured properly


### Getting Help

1. Check HTML reports for detailed error information
2. Review Exchange event logs on both environments
3. Consult Microsoft Exchange documentation
4. Open GitHub issues for toolkit-specific problems


## Contributing

We welcome contributions to improve this migration toolkit:

1. **Fork** the repository
2. **Create** a feature branch (`git checkout -b feature/improvement`)
3. **Test** thoroughly in lab environment
4. **Commit** your changes (`git commit -am 'Add new feature'`)
5. **Push** to the branch (`git push origin feature/improvement`)
6. **Create** a Pull Request


### Development Guidelines

- Follow PowerShell best practices
- Include comprehensive error handling
- Add HTML reporting for new features
- Update documentation for changes
- Test in multiple Exchange versions


## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- **Microsoft Exchange Team** for comprehensive migration documentation
- **PowerShell community** for best practices and examples
- **Exchange MVP community** for migration insights and troubleshooting guidance
- **Contributors** who have helped improve this toolkit


## Contact and Support

- **GitHub Issues**: Report bugs or request features
- **GitHub Discussions**: Community support and questions
- **Documentation**: Wiki pages for additional guidance


---

**⚠️ Important Disclaimer:** Always test in a lab environment before production use. Ensure proper backups are in place before starting any migration. This toolkit is provided as-is without warranty.

**Migration Toolkit Version:** 2.0**Last Updated:** January 2024**Compatibility:** Exchange 2013, 2016, 2019, 2022**PowerShell Version:** 5.1 or higher

The complete README.md file provides comprehensive documentation for the Exchange Cross-Forest Migration Toolkit, including step-by-step execution instructions, troubleshooting guides, and detailed parameter references. It clearly explains which scripts run on which environments and includes HTML reporting features with popup save dialogs for each script.
