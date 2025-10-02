# Bulk Litigation Hold Management System

## Overview

This system provides enterprise-grade bulk litigation hold management for large Microsoft 365 environments (120,000+ users). It efficiently identifies users with eligible licences and automatically enables litigation hold where appropriate, with comprehensive logging and safety controls.

## 🚀 Quick Start

### Prerequisites

1. **PowerShell 5.1+** with required modules:
   ```powershell
   Install-Module ExchangeOnlineManagement -Force
   Install-Module Microsoft.Graph -Force
   ```

2. **Authentication** - Connect to both services:
   ```powershell
   Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All"
   Connect-ExchangeOnline
   ```

### Basic Usage

```powershell
# Quick scan (preview mode)
.\Start-BulkLitigationHold.ps1 -QuickScan

# Test with limited users
.\Start-BulkLitigationHold.ps1 -TestRun

# Enable for all eligible users (with safety checks)
.\Start-BulkLitigationHold.ps1 -EnableAll -SafeMode
```

## 📁 File Structure

```
BulkLitigationHoldManager/
├── Set-BulkLitigationHold.ps1           # Main processing script
├── Start-BulkLitigationHold.ps1         # User-friendly launcher
├── BulkLitigationHoldUtilities.psm1     # Helper functions module
├── Config/
│   └── BulkLitigationHoldConfig.json    # License and processing configuration
├── Logs/                                # Generated logs and reports
└── README.md                           # This file
```

## 🏗️ Architecture

### High-Performance Design

- **Intelligent Batching**: Automatically optimises batch sizes based on environment size
- **Parallel Processing**: Uses PowerShell jobs for concurrent operations (up to 25 parallel jobs)
- **Memory Management**: Includes garbage collection for large datasets
- **API Optimisation**: Efficient Graph and Exchange Online API usage patterns

### Safety Features

- **Dry Run Mode**: Preview all changes before execution
- **Progressive Confirmation**: Multiple confirmation prompts for live operations
- **Error Handling**: Configurable error thresholds with graceful degradation
- **Rollback Capability**: Detailed logging enables manual rollback if needed
- **Safe Mode**: Reduced batch sizes and enhanced validation

## 🔍 License Detection

### Supported Licences for Litigation Hold

**Microsoft 365 Plans:**
- Microsoft 365 E3 (`ENTERPRISEPACK`, `SPE_E3`)
- Microsoft 365 E5 (`ENTERPRISEPREMIUM`, `SPE_E5`)
- Microsoft 365 Business Premium (`SPB`)
- Microsoft 365 F1 (`SPE_F1`)

**Exchange Plans:**
- Exchange Online Plan 2 (`EXCHANGEENTERPRISE`)
- Exchange Online Kiosk (`EXCHANGEDESKLESS`)
- Exchange Online Essentials (`EXCHANGEESSENTIALS`)

**Office 365 Plans:**
- Office 365 E1 (`STANDARDPACK`)
- Office 365 E3 (`ENTERPRISEPACKLRG`)
- Office 365 E4 (`ENTERPRISEWITHSCAL`)

**Government & Education:**
- Microsoft 365 GCC E3/E5 (`ENTERPRISEPACK_GOV`, `ENTERPRISEPREMIUM_GOV`)
- Education plans (`ENTERPRISEPACK_STUDENT`, etc.)

**Compliance Add-ons:**
- Microsoft 365 E5 Compliance (`INFORMATION_PROTECTION_COMPLIANCE`)

## ⚙️ Configuration

### Automatic Optimisation

The system automatically detects environment size and optimises settings:

| Environment Size | Batch Size | Concurrent Jobs | Memory Cleanup |
|-----------------|------------|----------------|----------------|
| < 1,000 users   | 100        | 5              | Every 10 batches |
| < 10,000 users  | 250        | 8              | Every 10 batches |
| < 50,000 users  | 500        | 10             | Every 10 batches |
| < 100,000 users | 750        | 15             | Every 10 batches |
| 100,000+ users  | 1,000      | 20             | Every 5 batches |

### Custom Configuration

Edit `Config/BulkLitigationHoldConfig.json` to customise:

```json
{
  "ProcessingConfiguration": {
    "Performance": {
      "DefaultBatchSize": 500,
      "RecommendedConcurrentJobs": 10
    },
    "Safety": {
      "DefaultMaxErrors": 100,
      "RequireConfirmationThreshold": 1000
    }
  }
}
```

## 📊 Comprehensive Logging

### Log Files Generated

1. **Main Log**: `BulkLitigationHold-YYYYMMDD-HHMMSS.log`
   - Complete execution log with timestamps
   - Progress tracking and status updates
   - Configuration and summary information

2. **Error Log**: `BulkLitigationHold-Errors-YYYYMMDD-HHMMSS.log`
   - Detailed error information
   - Failed operations with context
   - Troubleshooting information

3. **Detailed Report**: `BulkLitigationHold-Report-YYYYMMDD-HHMMSS.csv`
   - Complete user-level results
   - License information
   - Action taken for each user

4. **Performance Log**: `BulkLitigationHold-Performance-YYYYMMDD-HHMMSS.log`
   - System resource usage
   - Processing times and throughput
   - Memory and CPU utilisation

### Report Contents

Each report includes:
- User Principal Name and Display Name
- Current litigation hold status
- Eligible licences
- Action taken (Enabled/Already Enabled/Failed/etc.)
- Processing timestamp
- Error details (if applicable)

## 🛡️ Safety Controls

### Built-in Protections

1. **Prerequisite Validation**
   - PowerShell version check
   - Required module verification
   - Authentication status validation
   - System resource assessment

2. **Progressive Confirmation**
   - Initial parameter validation
   - Environment size warnings
   - Final execution confirmation
   - Type-to-confirm for live operations

3. **Error Management**
   - Configurable error thresholds
   - Graceful degradation on failures
   - Continue-on-error options
   - Detailed error logging

4. **Resource Protection**
   - Memory usage monitoring
   - Automatic garbage collection
   - API throttling respect
   - Concurrent job limiting

### Safe Mode Features

When using `-SafeMode`:
- Reduced batch size (100 users)
- Limited concurrent jobs (5)
- Enhanced error sensitivity (10 error limit)
- Additional validation steps
- More frequent progress reporting

## 📈 Performance Optimisation

### For Large Environments (100,000+ users)

**Recommended Settings:**
```powershell
.\Set-BulkLitigationHold.ps1 -BatchSize 1000 -MaxConcurrentJobs 20 -LogLevel Standard
```

**Best Practices:**
- Run during off-peak hours (2-6 AM)
- Ensure minimum 8GB available RAM
- Monitor network bandwidth usage
- Use Standard or Minimal log levels for better performance

### Memory Management

The system includes automatic memory management:
- Garbage collection every 5-10 batches
- Large object cleanup
- Progress reporting with memory usage
- Automatic batch size adjustment for low memory

## 🔧 Advanced Usage

### Custom Filtering

```powershell
# Process specific domain
.\Set-BulkLitigationHold.ps1 -UserFilter "*@contoso.com" -DryRun

# Process specific licence types
.\Set-BulkLitigationHold.ps1 -LicenseFilter "ENTERPRISEPACK,SPE_E5" -DryRun

# Continue despite errors
.\Set-BulkLitigationHold.ps1 -ContinueOnErrors -MaxErrors 500
```

### Integration with Automation

```powershell
# Scheduled execution
$Results = .\Set-BulkLitigationHold.ps1 -SkipConfirmation -LogLevel Minimal
if ($Results.Summary.TotalErrors -gt 0) {
    Send-MailMessage -To "admin@company.com" -Subject "Litigation Hold Errors" -Body $Results.Summary
}
```

### Performance Monitoring

```powershell
# Enable performance monitoring
Import-Module .\BulkLitigationHoldUtilities.psm1
$Monitor = Start-PerformanceMonitoring -LogPath ".\Logs\execution.log"
# ... run operations ...
$PerfSummary = Stop-PerformanceMonitoring -MonitoringJob $Monitor -LogPath ".\Logs\execution.log"
```

## 🚨 Troubleshooting

### Common Issues

**Authentication Failures:**
```powershell
# Re-authenticate
Disconnect-MgGraph
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All", "Organization.Read.All"

Disconnect-ExchangeOnline
Connect-ExchangeOnline
```

**Memory Issues (Large Environments):**
```powershell
# Reduce batch size and concurrent jobs
.\Set-BulkLitigationHold.ps1 -BatchSize 250 -MaxConcurrentJobs 5 -SafeMode
```

**API Throttling:**
```powershell
# Increase delays and reduce concurrency
.\Set-BulkLitigationHold.ps1 -MaxConcurrentJobs 3 -BatchSize 100
```

### Error Analysis

Check the error log for common patterns:
- **"Insufficient privileges"** → Check admin permissions
- **"Mailbox not found"** → User may lack Exchange Online license
- **"Throttling"** → Reduce concurrent jobs or add delays
- **"Memory"** → Use Safe Mode or reduce batch sizes

### Log Analysis

```powershell
# Quick error summary
Get-Content ".\Logs\*Errors*.log" | Select-String "ERROR" | Group-Object

# Performance analysis
Import-Csv ".\Logs\*Performance*.log" | Measure-Object CPU -Average
```

## 📋 Compliance & Auditing

### Audit Trail Features

- **Complete transaction logging** with timestamps
- **User-level action tracking** for compliance reporting
- **Before/after status recording** for audit purposes
- **Error correlation** with specific users and operations
- **Performance metrics** for operational reporting

### Compliance Reports

The system generates several compliance-focused reports:

1. **Executive Summary** (JSON format)
   - High-level statistics
   - Success/failure rates
   - Processing times
   - Resource utilisation

2. **Detailed Audit Trail** (CSV format)
   - User-by-user actions
   - License validation results
   - Timestamp information
   - Error details

3. **Compliance Status Report** (CSV format)
   - Current compliance state
   - Users requiring attention
   - License compliance status

## 🔄 Maintenance

### Log Cleanup

```powershell
# Clean up old logs (keeps 90 days by default)
Import-Module .\BulkLitigationHoldUtilities.psm1
Invoke-BulkLitigationHoldCleanup -LogPath ".\Logs" -RetentionDays 90 -CompressOldLogs
```

### Regular Health Checks

```powershell
# Weekly compliance validation
.\Start-BulkLitigationHold.ps1 -QuickScan

# Monthly full validation
.\Set-BulkLitigationHold.ps1 -DryRun -LogLevel Detailed
```

## 🎯 Use Cases

### Initial Deployment
1. **Assessment**: Run `-QuickScan` to understand scope
2. **Testing**: Use `-TestRun` with subset of users
3. **Staged Rollout**: Process by department or user pattern
4. **Full Deployment**: Use `-EnableAll` with appropriate safety settings

### Ongoing Maintenance
1. **New User Onboarding**: Regular scans for new eligible users
2. **License Changes**: Detection of users with newly assigned eligible licences
3. **Compliance Monitoring**: Regular reporting on litigation hold status
4. **Audit Preparation**: Generate comprehensive compliance reports

### Troubleshooting Scenarios
1. **Failed Operations**: Use error logs to identify and retry specific users
2. **Performance Issues**: Adjust batch sizes and concurrency based on system capacity
3. **License Changes**: Re-validate users when organisation licences change
4. **Rollback Requirements**: Use detailed logs to manually reverse changes if needed

---

## 📞 Support

For issues or questions:
1. Check the comprehensive error logs in the `Logs` folder
2. Review this documentation for configuration options
3. Use the `-DryRun` mode to test changes safely
4. Enable `-Verbose` logging for detailed troubleshooting information

This system is designed for enterprise reliability and scale while maintaining the flexibility needed for complex Microsoft 365 environments.