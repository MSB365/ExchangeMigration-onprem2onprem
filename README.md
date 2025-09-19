# Exchange Cross-Forest Migration Toolkit

A comprehensive PowerShell toolkit for automating Cross-Forest Exchange migrations between on-premises environments with Active Directory trust relationships.

## Overview

This toolkit splits the complex migration process into **6 numbered scripts** that must be executed in a specific order across **Source** and **Target** environments. Each script includes comprehensive HTML reporting with save dialog functionality.

## Architecture

\`\`\`
┌─────────────────┐    Trust Relationship    ┌─────────────────┐
│  SOURCE FOREST  │◄─────────────────────────►│  TARGET FOREST  │
│                 │                           │                 │
│ • Exchange Org  │                           │ • Exchange Org  │
│ • MRS-Proxy     │                           │ • Migration     │
│ • Source Users  │                           │   Endpoints     │
│                 │                           │ • Target DBs    │
└─────────────────┘                           └─────────────────┘
\`\`\`

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

## Prerequisites

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
\`\`\`csv
Identity
john.doe@source.tld
jane.smith@source.tld
admin@source.tld
\`\`\`

### users_migrate.csv (Step 3)
\`\`\`csv
EmailAddress
john.doe@source.tld
jane.smith@source.tld
admin@source.tld
\`\`\`

## Detailed Execution Guide

### Step 1: Source Environment Setup
**Environment:** SOURCE Exchange Management Shell  
**Script:** `01-Source-Environment-Setup.ps1`

```powershell
.\01-Source-Environment-Setup.ps1 `
    -SourceEwsFqdn "mail.source.tld" `
    -SourceDC "dc01.source.local" `
    -SourceAuthoritativeDomain "source.tld" `
    -TargetDC "dc01.target.local" `
    -TargetDeliveryDomain "target.tld"
