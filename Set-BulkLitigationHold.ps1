#Requires -Version 5.1
#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph

<#
.SYNOPSIS
    Bulk Litigation Hold Management for Large Organisations

.DESCRIPTION
    High-performance script designed to efficiently manage litigation hold settings across
    large environments (120,000+ users). Identifies licensed users eligible for litigation
    hold, compares against current status, and automatically enables where appropriate.
    
    Features:
    - Optimised for large-scale operations with intelligent batching
    - Parallel processing for maximum performance
    - Comprehensive logging and audit trail
    - Safe operation modes with dry-run and confirmation
    - Automatic license validation for litigation hold eligibility
    - Progress tracking and detailed reporting

.PARAMETER DryRun
    Preview mode - shows what would be changed without making modifications

.PARAMETER BatchSize
    Number of users to process in each batch (default: 500, optimal for large environments)

.PARAMETER MaxConcurrentJobs
    Maximum number of parallel processing jobs (default: 10)

.PARAMETER SkipConfirmation
    Skip interactive confirmation prompts (use with caution)

.PARAMETER LogLevel
    Logging verbosity: Minimal, Standard, Detailed, Verbose (default: Standard)

.PARAMETER OutputPath
    Custom path for logs and reports (default: .\Logs)

.PARAMETER LicenseFilter
    Comma-separated list of specific licence SKUs to process (optional)

.PARAMETER UserFilter
    UPN pattern to filter users (e.g., "*@contoso.com") - useful for testing

.PARAMETER ContinueOnErrors
    Continue processing even if some operations fail

.PARAMETER MaxErrors
    Maximum number of errors before stopping execution (default: 100)

.EXAMPLE
    .\Set-BulkLitigationHold.ps1 -DryRun
    Preview what changes would be made without executing them

.EXAMPLE
    .\Set-BulkLitigationHold.ps1 -BatchSize 1000 -MaxConcurrentJobs 15 -LogLevel Detailed
    Process with larger batches and more parallel jobs for maximum performance

.EXAMPLE
    .\Set-BulkLitigationHold.ps1 -UserFilter "*@contoso.com" -DryRun
    Test with a specific domain pattern first

.EXAMPLE
    .\Set-BulkLitigationHold.ps1 -LicenseFilter "ENTERPRISEPACK,SPE_E5" -SkipConfirmation
    Process only E3 and E5 licensed users without prompts

.NOTES
    Optimised for environments with 120,000+ users
    Requires Exchange Online Management and Microsoft Graph modules
    Uses efficient batching and parallel processing for performance
    All operations are logged for compliance and auditing
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$DryRun,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(50, 2000)]
    [int]$BatchSize = 500,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 25)]
    [int]$MaxConcurrentJobs = 10,
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipConfirmation,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Minimal", "Standard", "Detailed", "Verbose")]
    [string]$LogLevel = "Standard",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "",
    
    [Parameter(Mandatory = $false)]
    [string]$LicenseFilter = "",
    
    [Parameter(Mandatory = $false)]
    [string]$UserFilter = "",
    
    [Parameter(Mandatory = $false)]
    [switch]$ContinueOnErrors,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxErrors = 100
)

# Initialize script
$ScriptStart = Get-Date
$ScriptRoot = $PSScriptRoot
if (-not $OutputPath) {
    $OutputPath = Join-Path $ScriptRoot "Logs"
}

# Ensure output directory exists
if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

# Initialize logging
$LogFile = Join-Path $OutputPath "BulkLitigationHold-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
$ErrorLogFile = Join-Path $OutputPath "BulkLitigationHold-Errors-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
$ReportFile = Join-Path $OutputPath "BulkLitigationHold-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"

# Global counters and tracking
$Script:ProcessedUsers = 0
$Script:EligibleUsers = 0
$Script:AlreadyEnabled = 0
$Script:NewlyEnabled = 0
$Script:Errors = 0
$Script:SkippedUsers = 0

# Define litigation hold eligible licences
$LitigationHoldLicenses = @{
    # Exchange Plans
    "EXCHANGEENTERPRISE" = "Exchange Online Plan 2"
    "EXCHANGEDESKLESS" = "Exchange Online Kiosk"
    "EXCHANGEESSENTIALS" = "Exchange Online Essentials"
    
    # Microsoft 365 Plans
    "ENTERPRISEPACK" = "Microsoft 365 E3"
    "ENTERPRISEPREMIUM" = "Microsoft 365 E5"
    "SPE_E3" = "Microsoft 365 E3"
    "SPE_E5" = "Microsoft 365 E5"
    "SPB" = "Microsoft 365 Business Premium"
    "SPE_F1" = "Microsoft 365 F1"
    
    # Office 365 Plans
    "STANDARDPACK" = "Office 365 E1"
    "STANDARDWOFFPACK" = "Office 365 E2"
    "ENTERPRISEPACKLRG" = "Office 365 E3"
    "ENTERPRISEWITHSCAL" = "Office 365 E4"
    "STANDARDWOFFPACK_IW_STUDENT" = "Office 365 Education E1"
    
    # Compliance and Security Add-ons
    "INFORMATION_PROTECTION_COMPLIANCE" = "Microsoft 365 E5 Compliance"
    "THREAT_INTELLIGENCE" = "Microsoft 365 E5 Security"
    "ATP_ENTERPRISE" = "Microsoft Defender for Office 365"
    
    # Government Plans
    "ENTERPRISEPACK_GOV" = "Microsoft 365 GCC E3"
    "ENTERPRISEPREMIUM_GOV" = "Microsoft 365 GCC E5"
}

function Write-LogMessage {
    param(
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success", "Progress")]
        [string]$Type = "Info",
        [switch]$WriteToHost = $true
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Type] $Message"
    
    # Write to log file
    Add-Content -Path $LogFile -Value $LogEntry -Encoding UTF8
    
    # Write to error log if it's an error
    if ($Type -eq "Error") {
        Add-Content -Path $ErrorLogFile -Value $LogEntry -Encoding UTF8
    }
    
    # Write to console based on log level
    if ($WriteToHost) {
        $ShouldWrite = switch ($LogLevel) {
            "Minimal" { $Type -in @("Error", "Success") }
            "Standard" { $Type -in @("Info", "Warning", "Error", "Success") }
            "Detailed" { $Type -in @("Info", "Warning", "Error", "Success", "Progress") }
            "Verbose" { $true }
        }
        
        if ($ShouldWrite) {
            $Color = switch ($Type) {
                "Info" { "White" }
                "Warning" { "Yellow" }
                "Error" { "Red" }
                "Success" { "Green" }
                "Progress" { "Cyan" }
                default { "Gray" }
            }
            Write-Host $LogEntry -ForegroundColor $Color
        }
    }
}

function Test-LitigationHoldEligibility {
    param([array]$UserLicenses)
    
    $EligibleLicenses = @()
    
    foreach ($License in $UserLicenses) {
        $SkuId = $License.SkuId
        $SkuPartNumber = $License.SkuPartNumber
        
        # Check if this license supports litigation hold
        if ($LitigationHoldLicenses.ContainsKey($SkuPartNumber)) {
            $EligibleLicenses += @{
                SkuId = $SkuId
                SkuPartNumber = $SkuPartNumber
                DisplayName = $LitigationHoldLicenses[$SkuPartNumber]
            }
        }
    }
    
    return $EligibleLicenses
}

function Get-AllUsersWithLicenses {
    param([string]$Filter = "")
    
    Write-LogMessage "Starting comprehensive user and license data retrieval..." -Type Progress
    
    try {
        # Build Graph query with optimizations for large datasets
        $GraphQuery = @{
            All = $true
            Property = @('Id', 'UserPrincipalName', 'DisplayName', 'AccountEnabled', 'AssignedLicenses', 'UserType')
            PageSize = 999  # Maximum page size for Graph API
        }
        
        # Apply user filter if specified
        if ($Filter) {
            $GraphQuery.Filter = "startswith(userPrincipalName,'$($Filter.Replace('*',''))')"
        }
        
        Write-LogMessage "Retrieving all users from Microsoft Graph (this may take several minutes for large environments)..." -Type Progress
        $AllUsers = Get-MgUser @GraphQuery
        
        Write-LogMessage "Retrieved $($AllUsers.Count) total users from Microsoft Graph" -Type Success
        
        # Get license SKU mappings
        Write-LogMessage "Retrieving license SKU information..." -Type Progress
        $SkuList = Get-MgSubscribedSku
        $SkuMap = @{}
        
        foreach ($Sku in $SkuList) {
            $SkuMap[$Sku.SkuId] = $Sku.SkuPartNumber
        }
        
        Write-LogMessage "Processing license assignments for $($AllUsers.Count) users..." -Type Progress
        
        # Process users in batches to manage memory efficiently
        $ProcessedUsers = @()
        $BatchCounter = 0
        
        for ($i = 0; $i -lt $AllUsers.Count; $i += $BatchSize) {
            $BatchCounter++
            $Batch = $AllUsers[$i..([Math]::Min($i + $BatchSize - 1, $AllUsers.Count - 1))]
            
            Write-LogMessage "Processing batch $BatchCounter (users $($i + 1) to $([Math]::Min($i + $BatchSize, $AllUsers.Count)))" -Type Progress
            
            foreach ($User in $Batch) {
                if ($User.AssignedLicenses -and $User.AssignedLicenses.Count -gt 0 -and $User.AccountEnabled) {
                    # Map license SKU IDs to names
                    $UserLicenses = @()
                    foreach ($License in $User.AssignedLicenses) {
                        if ($SkuMap.ContainsKey($License.SkuId)) {
                            $UserLicenses += @{
                                SkuId = $License.SkuId
                                SkuPartNumber = $SkuMap[$License.SkuId]
                            }
                        }
                    }
                    
                    # Check if user has litigation hold eligible licenses
                    $EligibleLicenses = Test-LitigationHoldEligibility -UserLicenses $UserLicenses
                    
                    if ($EligibleLicenses.Count -gt 0) {
                        # Apply license filter if specified
                        if ($LicenseFilter) {
                            $FilteredLicenses = $LicenseFilter -split ","
                            $HasFilteredLicense = $false
                            
                            foreach ($EligibleLicense in $EligibleLicenses) {
                                if ($FilteredLicenses -contains $EligibleLicense.SkuPartNumber) {
                                    $HasFilteredLicense = $true
                                    break
                                }
                            }
                            
                            if (-not $HasFilteredLicense) {
                                continue
                            }
                        }
                        
                        $ProcessedUsers += @{
                            Id = $User.Id
                            UserPrincipalName = $User.UserPrincipalName
                            DisplayName = $User.DisplayName
                            EligibleLicenses = $EligibleLicenses
                            AllLicenses = $UserLicenses
                        }
                    }
                }
            }
            
            # Memory management for large datasets
            if ($BatchCounter % 10 -eq 0) {
                [System.GC]::Collect()
                Write-LogMessage "Processed $($ProcessedUsers.Count) eligible users so far (Memory cleanup performed)" -Type Progress
            }
        }
        
        Write-LogMessage "Identified $($ProcessedUsers.Count) users with litigation hold eligible licenses" -Type Success
        $Script:EligibleUsers = $ProcessedUsers.Count
        
        return $ProcessedUsers
        
    }
    catch {
        Write-LogMessage "Failed to retrieve users: $($_.Exception.Message)" -Type Error
        throw
    }
}

function Get-LitigationHoldStatus {
    param([array]$Users)
    
    Write-LogMessage "Retrieving current litigation hold status for $($Users.Count) users..." -Type Progress
    
    $UsersWithStatus = @()
    $BatchCounter = 0
    
    # Process in batches for efficiency
    for ($i = 0; $i -lt $Users.Count; $i += $BatchSize) {
        $BatchCounter++
        $Batch = $Users[$i..([Math]::Min($i + $BatchSize - 1, $Users.Count - 1))]
        
        Write-LogMessage "Checking litigation hold status for batch $BatchCounter (users $($i + 1) to $([Math]::Min($i + $BatchSize, $Users.Count)))" -Type Progress
        
        # Create array of UPNs for batch query
        $UPNList = $Batch | ForEach-Object { $_.UserPrincipalName }
        
        try {
            # Batch query for mailbox information
            $Mailboxes = Get-EXOMailbox -Identity $UPNList -Properties 'LitigationHoldEnabled','LitigationHoldDate','LitigationHoldOwner' -ErrorAction SilentlyContinue
            
            # Create lookup hashtable for faster processing
            $MailboxLookup = @{}
            foreach ($Mailbox in $Mailboxes) {
                $MailboxLookup[$Mailbox.UserPrincipalName] = $Mailbox
            }
            
            # Process each user in the batch
            foreach ($User in $Batch) {
                $MailboxInfo = $MailboxLookup[$User.UserPrincipalName]
                
                if ($MailboxInfo) {
                    $User.LitigationHoldEnabled = $MailboxInfo.LitigationHoldEnabled
                    $User.LitigationHoldDate = $MailboxInfo.LitigationHoldDate
                    $User.LitigationHoldOwner = $MailboxInfo.LitigationHoldOwner
                    $User.HasMailbox = $true
                } else {
                    $User.LitigationHoldEnabled = $false
                    $User.HasMailbox = $false
                    Write-LogMessage "No mailbox found for user: $($User.UserPrincipalName)" -Type Warning
                }
                
                $UsersWithStatus += $User
            }
            
        }
        catch {
            Write-LogMessage "Error processing batch $BatchCounter : $($_.Exception.Message)" -Type Error
            
            # Fall back to individual processing for this batch
            foreach ($User in $Batch) {
                try {
                    $Mailbox = Get-EXOMailbox -Identity $User.UserPrincipalName -Properties 'LitigationHoldEnabled','LitigationHoldDate','LitigationHoldOwner' -ErrorAction Stop
                    
                    $User.LitigationHoldEnabled = $Mailbox.LitigationHoldEnabled
                    $User.LitigationHoldDate = $Mailbox.LitigationHoldDate
                    $User.LitigationHoldOwner = $Mailbox.LitigationHoldOwner
                    $User.HasMailbox = $true
                    
                }
                catch {
                    $User.LitigationHoldEnabled = $false
                    $User.HasMailbox = $false
                    Write-LogMessage "Failed to get mailbox info for $($User.UserPrincipalName): $($_.Exception.Message)" -Type Error
                    $Script:Errors++
                }
                
                $UsersWithStatus += $User
            }
        }
        
        # Progress reporting
        $PercentComplete = [Math]::Round(($i / $Users.Count) * 100, 1)
        Write-LogMessage "Progress: $PercentComplete% complete ($($i + $Batch.Count)/$($Users.Count) users processed)" -Type Progress
        
        # Error threshold check
        if ($Script:Errors -gt $MaxErrors -and -not $ContinueOnErrors) {
            Write-LogMessage "Maximum error threshold ($MaxErrors) exceeded. Stopping execution." -Type Error
            throw "Too many errors encountered"
        }
    }
    
    # Calculate statistics
    $EnabledCount = ($UsersWithStatus | Where-Object { $_.LitigationHoldEnabled }).Count
    $DisabledCount = ($UsersWithStatus | Where-Object { -not $_.LitigationHoldEnabled -and $_.HasMailbox }).Count
    $NoMailboxCount = ($UsersWithStatus | Where-Object { -not $_.HasMailbox }).Count
    
    Write-LogMessage "Litigation hold status summary: $EnabledCount enabled, $DisabledCount disabled, $NoMailboxCount without mailbox" -Type Success
    
    $Script:AlreadyEnabled = $EnabledCount
    
    return $UsersWithStatus
}

function Set-LitigationHoldBulk {
    param([array]$UsersToEnable)
    
    if ($UsersToEnable.Count -eq 0) {
        Write-LogMessage "No users require litigation hold enablement" -Type Info
        return @()
    }
    
    Write-LogMessage "$(if ($DryRun) { '[DRY RUN] Would enable' } else { 'Enabling' }) litigation hold for $($UsersToEnable.Count) users..." -Type Progress
    
    $Results = @()
    $BatchCounter = 0
    $SuccessCount = 0
    $ErrorCount = 0
    
    # Process in smaller batches for bulk operations
    $BulkBatchSize = [Math]::Min($BatchSize, 100)  # Smaller batches for write operations
    
    for ($i = 0; $i -lt $UsersToEnable.Count; $i += $BulkBatchSize) {
        $BatchCounter++
        $Batch = $UsersToEnable[$i..([Math]::Min($i + $BulkBatchSize - 1, $UsersToEnable.Count - 1))]
        
        Write-LogMessage "$(if ($DryRun) { '[DRY RUN] ' } else { '' })Processing enablement batch $BatchCounter (users $($i + 1) to $([Math]::Min($i + $BulkBatchSize, $UsersToEnable.Count)))" -Type Progress
        
        if (-not $DryRun) {
            # Use parallel processing for bulk operations
            $Jobs = @()
            
            foreach ($User in $Batch) {
                $Job = Start-Job -ScriptBlock {
                    param($UserUPN, $LogFile)
                    
                    try {
                        Import-Module ExchangeOnlineManagement -Force
                        Set-Mailbox -Identity $UserUPN -LitigationHoldEnabled $true -ErrorAction Stop
                        
                        return @{
                            UserPrincipalName = $UserUPN
                            Success = $true
                            Error = $null
                            Timestamp = Get-Date
                        }
                    }
                    catch {
                        return @{
                            UserPrincipalName = $UserUPN
                            Success = $false
                            Error = $_.Exception.Message
                            Timestamp = Get-Date
                        }
                    }
                } -ArgumentList $User.UserPrincipalName, $LogFile
                
                $Jobs += $Job
                
                # Limit concurrent jobs
                while ((Get-Job -State Running).Count -ge $MaxConcurrentJobs) {
                    Start-Sleep -Milliseconds 100
                }
            }
            
            # Wait for all jobs in this batch to complete
            $Jobs | Wait-Job | Out-Null
            
            # Collect results
            foreach ($Job in $Jobs) {
                $Result = Receive-Job -Job $Job
                Remove-Job -Job $Job
                
                if ($Result.Success) {
                    Write-LogMessage "Successfully enabled litigation hold for: $($Result.UserPrincipalName)" -Type Success
                    $SuccessCount++
                } else {
                    Write-LogMessage "Failed to enable litigation hold for $($Result.UserPrincipalName): $($Result.Error)" -Type Error
                    $ErrorCount++
                    $Script:Errors++
                }
                
                $Results += $Result
            }
        } else {
            # Dry run - just log what would be done
            foreach ($User in $Batch) {
                Write-LogMessage "[DRY RUN] Would enable litigation hold for: $($User.UserPrincipalName) (Licenses: $($User.EligibleLicenses.DisplayName -join ', '))" -Type Info
                $Results += @{
                    UserPrincipalName = $User.UserPrincipalName
                    Success = $true
                    Error = $null
                    Timestamp = Get-Date
                    DryRun = $true
                }
            }
        }
        
        # Progress reporting
        $ProcessedSoFar = [Math]::Min($i + $BulkBatchSize, $UsersToEnable.Count)
        $PercentComplete = [Math]::Round(($ProcessedSoFar / $UsersToEnable.Count) * 100, 1)
        Write-LogMessage "Enablement progress: $PercentComplete% complete ($ProcessedSoFar/$($UsersToEnable.Count) users)" -Type Progress
        
        # Error threshold check
        if ($Script:Errors -gt $MaxErrors -and -not $ContinueOnErrors) {
            Write-LogMessage "Maximum error threshold ($MaxErrors) exceeded during enablement. Stopping." -Type Error
            break
        }
    }
    
    if (-not $DryRun) {
        Write-LogMessage "Bulk litigation hold enablement completed: $SuccessCount successful, $ErrorCount failed" -Type Success
        $Script:NewlyEnabled = $SuccessCount
    } else {
        Write-LogMessage "[DRY RUN] Would have enabled litigation hold for $($UsersToEnable.Count) users" -Type Success
    }
    
    return $Results
}

function Export-Results {
    param([array]$AllUsers, [array]$EnablementResults)
    
    Write-LogMessage "Generating comprehensive report..." -Type Progress
    
    $ReportData = @()
    
    foreach ($User in $AllUsers) {
        $EnablementResult = $EnablementResults | Where-Object { $_.UserPrincipalName -eq $User.UserPrincipalName }
        
        $ReportData += [PSCustomObject]@{
            UserPrincipalName = $User.UserPrincipalName
            DisplayName = $User.DisplayName
            HasMailbox = $User.HasMailbox
            LitigationHoldEnabled = $User.LitigationHoldEnabled
            LitigationHoldDate = $User.LitigationHoldDate
            LitigationHoldOwner = $User.LitigationHoldOwner
            EligibleLicenses = ($User.EligibleLicenses.DisplayName -join '; ')
            AllLicenses = ($User.AllLicenses.SkuPartNumber -join '; ')
            ActionTaken = if ($EnablementResult) { 
                if ($EnablementResult.DryRun) { "DRY RUN - Would Enable" }
                elseif ($EnablementResult.Success) { "Enabled" }
                else { "Failed to Enable" }
            } else { 
                if ($User.LitigationHoldEnabled) { "Already Enabled" }
                elseif (-not $User.HasMailbox) { "No Mailbox" }
                else { "No Action Required" }
            }
            Error = if ($EnablementResult -and -not $EnablementResult.Success) { $EnablementResult.Error } else { "" }
            ProcessedTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
    }
    
    # Export to CSV
    $ReportData | Export-Csv -Path $ReportFile -NoTypeInformation -Encoding UTF8
    Write-LogMessage "Detailed report exported to: $ReportFile" -Type Success
    
    # Generate summary statistics
    $Summary = @{
        TotalEligibleUsers = $AllUsers.Count
        AlreadyEnabled = ($ReportData | Where-Object { $_.LitigationHoldEnabled -eq $true }).Count
        NewlyEnabled = ($ReportData | Where-Object { $_.ActionTaken -eq "Enabled" }).Count
        Failed = ($ReportData | Where-Object { $_.ActionTaken -eq "Failed to Enable" }).Count
        NoMailbox = ($ReportData | Where-Object { $_.HasMailbox -eq $false }).Count
        TotalErrors = $Script:Errors
        ExecutionTime = (Get-Date) - $ScriptStart
    }
    
    return $Summary
}

# Main Execution
try {
    Write-Host "🏢 Bulk Litigation Hold Management System" -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
    Write-Host "Mode: $(if ($DryRun) { 'DRY RUN (Preview Only)' } else { 'LIVE EXECUTION' })" -ForegroundColor $(if ($DryRun) { 'Yellow' } else { 'Red' })
    Write-Host "Batch Size: $BatchSize users" -ForegroundColor Gray
    Write-Host "Max Concurrent Jobs: $MaxConcurrentJobs" -ForegroundColor Gray
    Write-Host "Log Level: $LogLevel" -ForegroundColor Gray
    
    Write-LogMessage "===== Bulk Litigation Hold Management Started =====" -Type Info
    Write-LogMessage "Script Version: 1.0 | Execution Mode: $(if ($DryRun) { 'DRY RUN' } else { 'LIVE' })" -Type Info
    Write-LogMessage "Configuration: BatchSize=$BatchSize, MaxJobs=$MaxConcurrentJobs, LogLevel=$LogLevel" -Type Info
    
    # Authentication check
    Write-LogMessage "Verifying authentication status..." -Type Progress
    
    try {
        $TestGraph = Get-MgContext -ErrorAction Stop
        if (-not $TestGraph) {
            throw "Not connected to Microsoft Graph"
        }
        Write-LogMessage "Microsoft Graph connection verified" -Type Success
    }
    catch {
        Write-LogMessage "Microsoft Graph not connected. Please run Connect-MgGraph first." -Type Error
        throw "Authentication required"
    }
    
    try {
        $TestExchange = Get-OrganizationConfig -ErrorAction Stop
        Write-LogMessage "Exchange Online connection verified" -Type Success
    }
    catch {
        Write-LogMessage "Exchange Online not connected. Please run Connect-ExchangeOnline first." -Type Error
        throw "Authentication required"
    }
    
    # Confirmation prompt
    if (-not $SkipConfirmation -and -not $DryRun) {
        Write-Host "`n⚠️  WARNING: This will modify litigation hold settings for eligible users" -ForegroundColor Yellow
        Write-Host "   This operation will scan all users and enable litigation hold where appropriate" -ForegroundColor Yellow
        $Confirmation = Read-Host "`nDo you want to continue? (Type 'YES' to confirm)"
        
        if ($Confirmation -ne 'YES') {
            Write-LogMessage "Operation cancelled by user" -Type Warning
            exit 0
        }
    }
    
    # Step 1: Get all users with eligible licenses
    Write-LogMessage "Phase 1: Identifying users with litigation hold eligible licenses..." -Type Progress
    $EligibleUsers = Get-AllUsersWithLicenses -Filter $UserFilter
    
    if ($EligibleUsers.Count -eq 0) {
        Write-LogMessage "No users found with litigation hold eligible licenses" -Type Warning
        exit 0
    }
    
    # Step 2: Check current litigation hold status
    Write-LogMessage "Phase 2: Checking current litigation hold status..." -Type Progress
    $UsersWithStatus = Get-LitigationHoldStatus -Users $EligibleUsers
    
    # Step 3: Identify users who need litigation hold enabled
    Write-LogMessage "Phase 3: Identifying users requiring litigation hold enablement..." -Type Progress
    $UsersToEnable = $UsersWithStatus | Where-Object { 
        -not $_.LitigationHoldEnabled -and $_.HasMailbox 
    }
    
    Write-LogMessage "Analysis complete: $($UsersToEnable.Count) users require litigation hold enablement" -Type Success
    
    # Step 4: Enable litigation hold for eligible users
    if ($UsersToEnable.Count -gt 0) {
        Write-LogMessage "Phase 4: $(if ($DryRun) { 'Simulating' } else { 'Executing' }) litigation hold enablement..." -Type Progress
        $EnablementResults = Set-LitigationHoldBulk -UsersToEnable $UsersToEnable
    } else {
        Write-LogMessage "Phase 4: No users require litigation hold enablement" -Type Info
        $EnablementResults = @()
    }
    
    # Step 5: Generate comprehensive report
    Write-LogMessage "Phase 5: Generating final report..." -Type Progress
    $Summary = Export-Results -AllUsers $UsersWithStatus -EnablementResults $EnablementResults
    
    # Final summary
    $ExecutionTime = (Get-Date) - $ScriptStart
    
    Write-Host "`n📊 Execution Summary" -ForegroundColor White -BackgroundColor DarkBlue
    Write-Host "Total Eligible Users: $($Summary.TotalEligibleUsers)" -ForegroundColor Cyan
    Write-Host "Already Enabled: $($Summary.AlreadyEnabled)" -ForegroundColor Green
    Write-Host "$(if ($DryRun) { 'Would Enable: ' } else { 'Newly Enabled: ' })$($Summary.NewlyEnabled)" -ForegroundColor $(if ($DryRun) { 'Yellow' } else { 'Green' })
    Write-Host "Failed Operations: $($Summary.Failed)" -ForegroundColor $(if ($Summary.Failed -gt 0) { 'Red' } else { 'Green' })
    Write-Host "Users Without Mailbox: $($Summary.NoMailbox)" -ForegroundColor Gray
    Write-Host "Total Errors: $($Summary.TotalErrors)" -ForegroundColor $(if ($Summary.TotalErrors -gt 0) { 'Red' } else { 'Green' })
    Write-Host "Execution Time: $($ExecutionTime.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
    Write-Host "Report File: $ReportFile" -ForegroundColor Cyan
    Write-Host "Log File: $LogFile" -ForegroundColor Cyan
    
    Write-LogMessage "===== Bulk Litigation Hold Management Completed Successfully =====" -Type Success
    Write-LogMessage "Final Summary: $($Summary.TotalEligibleUsers) eligible, $($Summary.NewlyEnabled) $(if ($DryRun) { 'would be ' } else { '' })enabled, $($Summary.TotalErrors) errors" -Type Success
    
    if ($DryRun) {
        Write-Host "`n💡 This was a DRY RUN. To execute changes, run again without -DryRun parameter" -ForegroundColor Yellow
    }
    
}
catch {
    $ExecutionTime = (Get-Date) - $ScriptStart
    Write-LogMessage "===== SCRIPT EXECUTION FAILED =====" -Type Error
    Write-LogMessage "Error: $($_.Exception.Message)" -Type Error
    Write-LogMessage "Execution time: $ExecutionTime" -Type Error
    
    Write-Host "`n💥 Script execution failed!" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Check log file for details: $LogFile" -ForegroundColor Yellow
    
    exit 1
}
finally {
    # Cleanup
    Write-LogMessage "Performing cleanup..." -Type Info
}