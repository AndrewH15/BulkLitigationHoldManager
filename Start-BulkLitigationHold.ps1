#Requires -Version 5.1

<#
.SYNOPSIS
    Quick launcher for bulk litigation hold management operations

.DESCRIPTION
    Simplified interface for common bulk litigation hold scenarios with built-in safety checks.
    Automatically detects environment size and applies optimal configurations.

.PARAMETER QuickScan
    Perform a quick scan to identify users requiring litigation hold (dry run mode)

.PARAMETER EnableAll
    Enable litigation hold for all eligible users (requires confirmation)

.PARAMETER TestRun
    Run with a small subset of users for testing (first 100 eligible users)

.PARAMETER SafeMode
    Run with enhanced safety checks and smaller batch sizes

.PARAMETER ConfigPath
    Path to custom configuration file

.EXAMPLE
    .\Start-BulkLitigationHold.ps1 -QuickScan
    Preview users who would be affected by litigation hold enablement

.EXAMPLE  
    .\Start-BulkLitigationHold.ps1 -TestRun
    Test with first 100 eligible users

.EXAMPLE
    .\Start-BulkLitigationHold.ps1 -EnableAll -SafeMode
    Enable litigation hold for all eligible users with enhanced safety

.NOTES
    This script automatically detects environment size and optimises settings
    Always performs prerequisite checks before execution
#>

[CmdletBinding(DefaultParameterSetName = "QuickScan")]
param(
    [Parameter(ParameterSetName = "QuickScan")]
    [switch]$QuickScan,
    
    [Parameter(ParameterSetName = "EnableAll")]
    [switch]$EnableAll,
    
    [Parameter(ParameterSetName = "TestRun")]
    [switch]$TestRun,
    
    [Parameter(Mandatory = $false)]
    [switch]$SafeMode,
    
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = ""
)

# Script configuration
$ScriptRoot = $PSScriptRoot
$MainScript = Join-Path $ScriptRoot "Set-BulkLitigationHold.ps1"
$UtilitiesModule = Join-Path $ScriptRoot "BulkLitigationHoldUtilities.psm1"
$DefaultConfigPath = Join-Path $ScriptRoot "Config\BulkLitigationHoldConfig.json"

Write-Host "🏢 Bulk Litigation Hold Management Launcher" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan

# Validate prerequisites
if (-not (Test-Path $MainScript)) {
    Write-Host "❌ Main script not found: $MainScript" -ForegroundColor Red
    exit 1
}

# Import utilities if available
if (Test-Path $UtilitiesModule) {
    try {
        Import-Module $UtilitiesModule -Force
        Write-Host "✅ Utilities module loaded" -ForegroundColor Green
    }
    catch {
        Write-Host "⚠️ Could not load utilities module: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Load configuration
if (-not $ConfigPath) {
    $ConfigPath = $DefaultConfigPath
}

$Config = $null
if (Test-Path $ConfigPath) {
    try {
        $Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        Write-Host "✅ Configuration loaded from: $ConfigPath" -ForegroundColor Green
    }
    catch {
        Write-Host "⚠️ Could not load configuration: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Perform prerequisite checks
Write-Host "`n🔍 Performing prerequisite checks..." -ForegroundColor Yellow

if (Get-Command Test-BulkLitigationHoldPrerequisites -ErrorAction SilentlyContinue) {
    $Prerequisites = Test-BulkLitigationHoldPrerequisites
    
    # Display prerequisite results
    Write-Host "   PowerShell Version: $(if ($Prerequisites.PowerShellVersion) { '✅' } else { '❌' })" -ForegroundColor $(if ($Prerequisites.PowerShellVersion) { 'Green' } else { 'Red' })
    Write-Host "   Exchange Online Module: $(if ($Prerequisites.RequiredModules.ExchangeOnlineManagement) { '✅' } else { '❌' })" -ForegroundColor $(if ($Prerequisites.RequiredModules.ExchangeOnlineManagement) { 'Green' } else { 'Red' })
    Write-Host "   Microsoft Graph Module: $(if ($Prerequisites.RequiredModules.MicrosoftGraph) { '✅' } else { '❌' })" -ForegroundColor $(if ($Prerequisites.RequiredModules.MicrosoftGraph) { 'Green' } else { 'Red' })
    Write-Host "   Graph Connection: $(if ($Prerequisites.Authentication.GraphConnection) { '✅' } else { '❌' })" -ForegroundColor $(if ($Prerequisites.Authentication.GraphConnection) { 'Green' } else { 'Red' })
    Write-Host "   Exchange Connection: $(if ($Prerequisites.Authentication.ExchangeConnection) { '✅' } else { '❌' })" -ForegroundColor $(if ($Prerequisites.Authentication.ExchangeConnection) { 'Green' } else { 'Red' })
    Write-Host "   Available Memory: $($Prerequisites.SystemResources.AvailableMemory) GB" -ForegroundColor Gray
    Write-Host "   Available Disk Space: $($Prerequisites.SystemResources.DiskSpace) GB" -ForegroundColor Gray
    
    # Check for critical prerequisites
    $CriticalIssues = @()
    if (-not $Prerequisites.PowerShellVersion) { $CriticalIssues += "PowerShell 5.1+ required" }
    if (-not $Prerequisites.RequiredModules.ExchangeOnlineManagement) { $CriticalIssues += "ExchangeOnlineManagement module not installed" }
    if (-not $Prerequisites.RequiredModules.MicrosoftGraph) { $CriticalIssues += "Microsoft.Graph module not installed" }
    if (-not $Prerequisites.Authentication.GraphConnection) { $CriticalIssues += "Not connected to Microsoft Graph" }
    if (-not $Prerequisites.Authentication.ExchangeConnection) { $CriticalIssues += "Not connected to Exchange Online" }
    
    if ($CriticalIssues.Count -gt 0) {
        Write-Host "`n❌ Critical Prerequisites Missing:" -ForegroundColor Red
        $CriticalIssues | ForEach-Object { Write-Host "   • $_" -ForegroundColor Red }
        Write-Host "`nPlease resolve these issues before continuing." -ForegroundColor Yellow
        exit 1
    }
} else {
    Write-Host "   ⚠️ Prerequisite checker not available - performing basic checks" -ForegroundColor Yellow
    
    # Basic checks
    try {
        $GraphTest = Get-MgContext -ErrorAction Stop
        Write-Host "   Graph Connection: ✅" -ForegroundColor Green
    }
    catch {
        Write-Host "   Graph Connection: ❌ Please run Connect-MgGraph" -ForegroundColor Red
    }
    
    try {
        $ExchangeTest = Get-OrganizationConfig -ErrorAction Stop
        Write-Host "   Exchange Connection: ✅" -ForegroundColor Green
    }
    catch {
        Write-Host "   Exchange Connection: ❌ Please run Connect-ExchangeOnline" -ForegroundColor Red
    }
}

# Build parameters for main script
$ScriptParams = @{}

# Configure based on selected mode
switch ($PSCmdlet.ParameterSetName) {
    "QuickScan" {
        Write-Host "`n📋 Quick Scan Mode: Preview users requiring litigation hold" -ForegroundColor Yellow
        $ScriptParams.DryRun = $true
        $ScriptParams.LogLevel = "Standard"
    }
    
    "EnableAll" {
        Write-Host "`n⚡ Enable All Mode: Process all eligible users" -ForegroundColor Yellow
        $ScriptParams.LogLevel = "Detailed"
        
        if (-not $SafeMode) {
            Write-Host "   ⚠️ Running in full execution mode" -ForegroundColor Yellow
        }
    }
    
    "TestRun" {
        Write-Host "`n🧪 Test Run Mode: Limited to first 100 eligible users" -ForegroundColor Yellow
        $ScriptParams.DryRun = $true
        $ScriptParams.UserFilter = "*"  # Process all domains but limit will be applied in script
        $ScriptParams.LogLevel = "Detailed"
    }
}

# Apply safe mode settings
if ($SafeMode) {
    Write-Host "🛡️ Safe Mode: Enhanced safety checks enabled" -ForegroundColor Green
    $ScriptParams.BatchSize = 100
    $ScriptParams.MaxConcurrentJobs = 5
    $ScriptParams.ContinueOnErrors = $false
    $ScriptParams.MaxErrors = 10
}

# Auto-detect environment size and optimise if utilities are available
if (Get-Command Get-OptimalBatchConfiguration -ErrorAction SilentlyContinue) {
    try {
        Write-Host "`n⚙️ Auto-detecting environment size for optimisation..." -ForegroundColor Yellow
        
        # Quick user count estimation
        $SampleUsers = Get-MgUser -Top 1000 -Property UserPrincipalName
        $EstimatedTotalUsers = if ($SampleUsers.Count -eq 1000) {
            # Likely more than 1000 users, get organization stats if possible
            try {
                $OrgStats = Get-MgOrganization
                if ($OrgStats.DirectorySizeQuota.Used) {
                    [int]$OrgStats.DirectorySizeQuota.Used
                } else {
                    50000  # Conservative estimate for large organizations
                }
            }
            catch {
                25000  # Default estimate if org stats unavailable
            }
        } else {
            $SampleUsers.Count
        }
        
        Write-Host "   Estimated environment size: $EstimatedTotalUsers users" -ForegroundColor Gray
        
        $OptimalConfig = Get-OptimalBatchConfiguration -TotalUsers $EstimatedTotalUsers
        
        # Apply optimal configuration unless overridden by safe mode
        if (-not $SafeMode) {
            $ScriptParams.BatchSize = $OptimalConfig.BatchSize
            $ScriptParams.MaxConcurrentJobs = $OptimalConfig.ConcurrentJobs
        }
        
        Write-Host "   Optimal batch size: $($ScriptParams.BatchSize)" -ForegroundColor Gray
        Write-Host "   Optimal concurrent jobs: $($ScriptParams.MaxConcurrentJobs)" -ForegroundColor Gray
        
        if ($OptimalConfig.Warnings.Count -gt 0) {
            Write-Host "   Recommendations:" -ForegroundColor Yellow
            $OptimalConfig.Warnings | ForEach-Object { Write-Host "     • $_" -ForegroundColor Yellow }
        }
        
    }
    catch {
        Write-Host "   ⚠️ Could not auto-detect environment size: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Display final configuration
Write-Host "`n📊 Execution Configuration:" -ForegroundColor White
Write-Host "   Mode: $(if ($ScriptParams.DryRun) { 'DRY RUN (Preview)' } else { 'LIVE EXECUTION' })" -ForegroundColor $(if ($ScriptParams.DryRun) { 'Yellow' } else { 'Red' })
if ($ScriptParams.BatchSize) { Write-Host "   Batch Size: $($ScriptParams.BatchSize)" -ForegroundColor Gray }
if ($ScriptParams.MaxConcurrentJobs) { Write-Host "   Concurrent Jobs: $($ScriptParams.MaxConcurrentJobs)" -ForegroundColor Gray }
Write-Host "   Log Level: $($ScriptParams.LogLevel)" -ForegroundColor Gray
Write-Host "   Safe Mode: $(if ($SafeMode) { 'Enabled' } else { 'Disabled' })" -ForegroundColor $(if ($SafeMode) { 'Green' } else { 'Yellow' })

# Final confirmation for live execution
if (-not $ScriptParams.DryRun -and -not $TestRun) {
    Write-Host "`n⚠️  FINAL CONFIRMATION" -ForegroundColor Red -BackgroundColor Yellow
    Write-Host "This will enable litigation hold for eligible users in your organisation." -ForegroundColor Yellow
    Write-Host "This operation will modify user mailbox settings." -ForegroundColor Yellow
    
    $FinalConfirmation = Read-Host "`nType 'EXECUTE' to proceed with live changes"
    
    if ($FinalConfirmation -ne 'EXECUTE') {
        Write-Host "❌ Operation cancelled" -ForegroundColor Yellow
        exit 0
    }
    
    $ScriptParams.SkipConfirmation = $true
}

# Execute main script
try {
    Write-Host "`n🚀 Starting bulk litigation hold management..." -ForegroundColor Green
    
    & $MainScript @ScriptParams
    
    $ExitCode = $LASTEXITCODE
    
    if ($ExitCode -eq 0 -or $ExitCode -eq $null) {
        Write-Host "`n✅ Operation completed successfully!" -ForegroundColor Green
        
        if ($ScriptParams.DryRun) {
            Write-Host "💡 This was a preview. To execute changes, use:" -ForegroundColor Yellow
            Write-Host "   .\Start-BulkLitigationHold.ps1 -EnableAll" -ForegroundColor Cyan
        }
    } else {
        Write-Host "`n❌ Operation completed with errors (Exit Code: $ExitCode)" -ForegroundColor Red
    }
    
} catch {
    Write-Host "`n💥 Operation failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Quick actions menu
Write-Host "`n📈 Quick Actions:" -ForegroundColor White
Write-Host "   • View logs: explorer `"$ScriptRoot\Logs`"" -ForegroundColor Gray
Write-Host "   • Re-run scan: .\Start-BulkLitigationHold.ps1 -QuickScan" -ForegroundColor Gray
Write-Host "   • Test run: .\Start-BulkLitigationHold.ps1 -TestRun" -ForegroundColor Gray
if ($ScriptParams.DryRun) {
    Write-Host "   • Execute changes: .\Start-BulkLitigationHold.ps1 -EnableAll" -ForegroundColor Cyan
}