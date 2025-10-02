#Requires -Version 5.1

<#
.SYNOPSIS
    Utility functions for bulk litigation hold management operations

.DESCRIPTION
    Supporting functions and utilities for the bulk litigation hold management system.
    Provides performance monitoring, validation helpers, and maintenance operations.

.NOTES
    This module contains helper functions used by Set-BulkLitigationHold.ps1
    Optimised for large-scale enterprise environments (120,000+ users)
#>

function Test-BulkLitigationHoldPrerequisites {
    <#
    .SYNOPSIS
        Validates system prerequisites for bulk litigation hold operations
    #>
    [CmdletBinding()]
    param()
    
    $Prerequisites = @{
        PowerShellVersion = $false
        RequiredModules = @{
            ExchangeOnlineManagement = $false
            MicrosoftGraph = $false
        }
        Authentication = @{
            GraphConnection = $false
            ExchangeConnection = $false
        }
        SystemResources = @{
            AvailableMemory = 0
            DiskSpace = 0
        }
        Permissions = @{
            GraphPermissions = @()
            ExchangePermissions = @()
        }
    }
    
    # Check PowerShell version
    if ($PSVersionTable.PSVersion.Major -ge 5) {
        $Prerequisites.PowerShellVersion = $true
    }
    
    # Check required modules
    try {
        $ExchangeModule = Get-Module -Name ExchangeOnlineManagement -ListAvailable | Select-Object -First 1
        if ($ExchangeModule) {
            $Prerequisites.RequiredModules.ExchangeOnlineManagement = $true
        }
        
        $GraphModule = Get-Module -Name Microsoft.Graph -ListAvailable | Select-Object -First 1
        if ($GraphModule) {
            $Prerequisites.RequiredModules.MicrosoftGraph = $true
        }
    }
    catch {
        Write-Warning "Error checking modules: $($_.Exception.Message)"
    }
    
    # Check authentication status
    try {
        $GraphContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($GraphContext) {
            $Prerequisites.Authentication.GraphConnection = $true
        }
        
        $ExchangeSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
        if ($ExchangeSession) {
            $Prerequisites.Authentication.ExchangeConnection = $true
        }
    }
    catch {
        Write-Warning "Error checking authentication: $($_.Exception.Message)"
    }
    
    # Check system resources
    try {
        $Memory = Get-CimInstance -ClassName Win32_OperatingSystem
        $Prerequisites.SystemResources.AvailableMemory = [Math]::Round($Memory.FreePhysicalMemory / 1MB, 2)
        
        $Disk = Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | Select-Object -First 1
        $Prerequisites.SystemResources.DiskSpace = [Math]::Round($Disk.FreeSpace / 1GB, 2)
    }
    catch {
        Write-Warning "Error checking system resources: $($_.Exception.Message)"
    }
    
    return $Prerequisites
}

function Get-OptimalBatchConfiguration {
    <#
    .SYNOPSIS
        Calculates optimal batch sizes and concurrency settings based on environment size and system resources
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$TotalUsers,
        
        [Parameter(Mandatory = $false)]
        [int]$AvailableMemoryMB = 4096,
        
        [Parameter(Mandatory = $false)]
        [int]$NetworkBandwidthMbps = 100
    )
    
    $Configuration = @{
        BatchSize = 500
        ConcurrentJobs = 10
        MemoryCleanupInterval = 10
        ThrottlingDelay = 100
        RecommendedRuntime = "Off-Hours"
        Warnings = @()
    }
    
    # Adjust based on environment size
    switch ($TotalUsers) {
        { $_ -lt 1000 } {
            $Configuration.BatchSize = 100
            $Configuration.ConcurrentJobs = 5
        }
        { $_ -lt 10000 } {
            $Configuration.BatchSize = 250
            $Configuration.ConcurrentJobs = 8
        }
        { $_ -lt 50000 } {
            $Configuration.BatchSize = 500
            $Configuration.ConcurrentJobs = 10
        }
        { $_ -lt 100000 } {
            $Configuration.BatchSize = 750
            $Configuration.ConcurrentJobs = 15
        }
        { $_ -ge 100000 } {
            $Configuration.BatchSize = 1000
            $Configuration.ConcurrentJobs = 20
            $Configuration.MemoryCleanupInterval = 5
            $Configuration.RecommendedRuntime = "Off-Hours (2-6 AM)"
            $Configuration.Warnings += "Large environment detected. Consider running during off-peak hours."
        }
    }
    
    # Adjust based on available memory
    if ($AvailableMemoryMB -lt 2048) {
        $Configuration.BatchSize = [Math]::Max(50, [Math]::Floor($Configuration.BatchSize * 0.5))
        $Configuration.ConcurrentJobs = [Math]::Max(2, [Math]::Floor($Configuration.ConcurrentJobs * 0.5))
        $Configuration.Warnings += "Low memory detected. Reducing batch size and concurrency."
    }
    elseif ($AvailableMemoryMB -gt 8192) {
        $Configuration.BatchSize = [Math]::Min(2000, [Math]::Floor($Configuration.BatchSize * 1.5))
        $Configuration.ConcurrentJobs = [Math]::Min(25, [Math]::Floor($Configuration.ConcurrentJobs * 1.5))
    }
    
    # Adjust based on network bandwidth
    if ($NetworkBandwidthMbps -lt 50) {
        $Configuration.ThrottlingDelay = 500
        $Configuration.ConcurrentJobs = [Math]::Max(2, [Math]::Floor($Configuration.ConcurrentJobs * 0.7))
        $Configuration.Warnings += "Limited bandwidth detected. Increasing throttling delays."
    }
    
    return $Configuration
}

function Start-PerformanceMonitoring {
    <#
    .SYNOPSIS
        Initiates performance monitoring for bulk operations
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )
    
    $MonitoringJob = Start-Job -ScriptBlock {
        param($LogPath)
        
        $MonitoringLog = $LogPath -replace '\.log$', '-Performance.log'
        $StartTime = Get-Date
        
        while ($true) {
            try {
                $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                
                # Get system performance metrics
                $CPU = Get-CimInstance -ClassName Win32_Processor | Measure-Object -Property LoadPercentage -Average | Select-Object -ExpandProperty Average
                $Memory = Get-CimInstance -ClassName Win32_OperatingSystem
                $MemoryUsage = [Math]::Round((($Memory.TotalVisibleMemorySize - $Memory.FreePhysicalMemory) / $Memory.TotalVisibleMemorySize) * 100, 2)
                
                # Get PowerShell process metrics
                $PSProcess = Get-Process -Name PowerShell | Where-Object { $_.Id -eq $PID }
                $PSMemoryMB = [Math]::Round($PSProcess.WorkingSet64 / 1MB, 2)
                
                # Get network activity (simplified)
                $NetworkCounters = Get-CimInstance -ClassName Win32_PerfRawData_Tcpip_NetworkInterface | Where-Object { $_.Name -notlike "*Loopback*" -and $_.Name -notlike "*Teredo*" } | Select-Object -First 1
                
                $LogEntry = "$Timestamp,CPU:$CPU%,Memory:$MemoryUsage%,PSMemory:${PSMemoryMB}MB"
                Add-Content -Path $MonitoringLog -Value $LogEntry
                
                Start-Sleep -Seconds 30
            }
            catch {
                # Silently continue if monitoring fails
                Start-Sleep -Seconds 30
            }
        }
    } -ArgumentList $LogPath
    
    return $MonitoringJob
}

function Stop-PerformanceMonitoring {
    <#
    .SYNOPSIS
        Stops performance monitoring and generates summary
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Job]$MonitoringJob,
        
        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )
    
    try {
        $MonitoringJob | Stop-Job -ErrorAction SilentlyContinue
        $MonitoringJob | Remove-Job -ErrorAction SilentlyContinue
        
        $MonitoringLog = $LogPath -replace '\.log$', '-Performance.log'
        
        if (Test-Path $MonitoringLog) {
            $PerformanceData = Import-Csv -Path $MonitoringLog -Header @('Timestamp', 'CPU', 'Memory', 'PSMemory')
            
            # Calculate averages
            $AvgCPU = ($PerformanceData | ForEach-Object { [double]($_.CPU -replace '[^\d.]', '') } | Measure-Object -Average).Average
            $AvgMemory = ($PerformanceData | ForEach-Object { [double]($_.Memory -replace '[^\d.]', '') } | Measure-Object -Average).Average
            $MaxPSMemory = ($PerformanceData | ForEach-Object { [double]($_.PSMemory -replace '[^\d.]', '') } | Measure-Object -Maximum).Maximum
            
            $Summary = @{
                AverageCPU = [Math]::Round($AvgCPU, 2)
                AverageMemory = [Math]::Round($AvgMemory, 2)
                MaxPowerShellMemory = [Math]::Round($MaxPSMemory, 2)
                DataPoints = $PerformanceData.Count
            }
            
            return $Summary
        }
    }
    catch {
        Write-Warning "Error stopping performance monitoring: $($_.Exception.Message)"
    }
    
    return $null
}

function Test-LitigationHoldLicense {
    <#
    .SYNOPSIS
        Enhanced license validation with detailed license information
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$AssignedLicenses,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$SkuMap = @{},
        
        [Parameter(Mandatory = $false)]
        [string]$ConfigPath = ""
    )
    
    # Load configuration if provided - Only licenses with native litigation hold support
    $LitigationHoldLicenses = @{
        "EXCHANGEENTERPRISE" = "Exchange Online Plan 2"
        "ENTERPRISEPACK" = "Microsoft 365 E3"
        "ENTERPRISEPREMIUM" = "Microsoft 365 E5"
        "SPE_E3" = "Microsoft 365 E3"
        "SPE_E5" = "Microsoft 365 E5"
        "SPB" = "Microsoft 365 Business Premium"
    }
    
    if ($ConfigPath -and (Test-Path $ConfigPath)) {
        try {
            $Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
            $ConfigLicenses = @{}
            
            foreach ($Category in $Config.LicenseConfiguration.LitigationHoldEligibleLicenses.PSObject.Properties) {
                foreach ($License in $Category.Value.PSObject.Properties) {
                    if ($License.Value.LitigationHoldSupported) {
                        $ConfigLicenses[$License.Name] = $License.Value.DisplayName
                    }
                }
            }
            
            if ($ConfigLicenses.Count -gt 0) {
                $LitigationHoldLicenses = $ConfigLicenses
            }
        }
        catch {
            Write-Warning "Could not load license configuration from $ConfigPath : $($_.Exception.Message)"
        }
    }
    
    $Result = @{
        IsEligible = $false
        EligibleLicenses = @()
        AllLicenses = @()
        HighestPriority = "None"
    }
    
    foreach ($License in $AssignedLicenses) {
        $SkuPartNumber = if ($SkuMap.ContainsKey($License.SkuId)) { 
            $SkuMap[$License.SkuId] 
        } else { 
            $License.SkuId 
        }
        
        $Result.AllLicenses += @{
            SkuId = $License.SkuId
            SkuPartNumber = $SkuPartNumber
        }
        
        if ($LitigationHoldLicenses.ContainsKey($SkuPartNumber)) {
            $Result.IsEligible = $true
            $Result.EligibleLicenses += @{
                SkuId = $License.SkuId
                SkuPartNumber = $SkuPartNumber
                DisplayName = $LitigationHoldLicenses[$SkuPartNumber]
            }
            
            # Determine priority (simplified logic)
            $Priority = switch ($SkuPartNumber) {
                { $_ -like "*E5*" } { "Very High" }
                { $_ -like "*E3*" } { "High" }
                { $_ -like "*ENTERPRISE*" } { "High" }
                default { "Medium" }
            }
            
            if ($Result.HighestPriority -eq "None" -or 
                ($Priority -eq "Very High") -or
                ($Priority -eq "High" -and $Result.HighestPriority -ne "Very High")) {
                $Result.HighestPriority = $Priority
            }
        }
    }
    
    return $Result
}

function Export-BulkOperationReport {
    <#
    .SYNOPSIS
        Generates comprehensive reports for bulk litigation hold operations
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$ProcessedUsers,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$ExecutionSummary,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$PerformanceMetrics = @{}
    )
    
    $ReportTimestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    
    # Main detailed report
    $DetailedReportPath = Join-Path $OutputPath "DetailedReport-$ReportTimestamp.csv"
    $ProcessedUsers | Export-Csv -Path $DetailedReportPath -NoTypeInformation -Encoding UTF8
    
    # Executive summary report
    $SummaryReportPath = Join-Path $OutputPath "ExecutiveSummary-$ReportTimestamp.json"
    $ExecutionSummary | ConvertTo-Json -Depth 10 | Out-File -FilePath $SummaryReportPath -Encoding UTF8
    
    # Performance report (if available)
    if ($PerformanceMetrics.Count -gt 0) {
        $PerformanceReportPath = Join-Path $OutputPath "PerformanceReport-$ReportTimestamp.json"
        $PerformanceMetrics | ConvertTo-Json -Depth 5 | Out-File -FilePath $PerformanceReportPath -Encoding UTF8
    }
    
    # Compliance audit report
    $ComplianceReportPath = Join-Path $OutputPath "ComplianceAudit-$ReportTimestamp.csv"
    $ComplianceData = $ProcessedUsers | Select-Object UserPrincipalName, DisplayName, 
        @{Name="ComplianceStatus"; Expression={
            if ($_.LitigationHoldEnabled) { "Compliant" }
            elseif ($_.HasMailbox -and $_.ActionTaken -eq "Enabled") { "Newly Compliant" }
            elseif (-not $_.HasMailbox) { "No Mailbox - N/A" }
            else { "Non-Compliant" }
        }},
        @{Name="LicenseCompliance"; Expression={
            if ($_.EligibleLicenses) { "Has Eligible License" } else { "No Eligible License" }
        }},
        LitigationHoldEnabled, ActionTaken, ProcessedTimestamp
    
    $ComplianceData | Export-Csv -Path $ComplianceReportPath -NoTypeInformation -Encoding UTF8
    
    return @{
        DetailedReport = $DetailedReportPath
        SummaryReport = $SummaryReportPath
        ComplianceReport = $ComplianceReportPath
        PerformanceReport = if ($PerformanceMetrics.Count -gt 0) { $PerformanceReportPath } else { $null }
    }
}

function Invoke-BulkLitigationHoldCleanup {
    <#
    .SYNOPSIS
        Performs cleanup operations after bulk litigation hold processing
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogPath,
        
        [Parameter(Mandatory = $false)]
        [int]$RetentionDays = 90,
        
        [Parameter(Mandatory = $false)]
        [switch]$CompressOldLogs
    )
    
    try {
        $LogDirectory = Split-Path $LogPath -Parent
        $CutoffDate = (Get-Date).AddDays(-$RetentionDays)
        
        # Find old log files
        $OldLogs = Get-ChildItem -Path $LogDirectory -Filter "*.log" | Where-Object { $_.LastWriteTime -lt $CutoffDate }
        
        if ($OldLogs.Count -gt 0) {
            Write-Host "Found $($OldLogs.Count) log files older than $RetentionDays days"
            
            if ($CompressOldLogs) {
                # Compress old logs before deletion
                foreach ($Log in $OldLogs) {
                    try {
                        $ZipPath = $Log.FullName -replace '\.log$', '.zip'
                        Compress-Archive -Path $Log.FullName -DestinationPath $ZipPath -Force
                        Remove-Item -Path $Log.FullName -Force
                        Write-Host "Compressed and removed: $($Log.Name)"
                    }
                    catch {
                        Write-Warning "Failed to compress $($Log.Name): $($_.Exception.Message)"
                    }
                }
            } else {
                # Delete old logs
                $OldLogs | Remove-Item -Force
                Write-Host "Removed $($OldLogs.Count) old log files"
            }
        }
        
        # Clean up temporary files
        $TempFiles = Get-ChildItem -Path $LogDirectory -Filter "*.tmp" -ErrorAction SilentlyContinue
        if ($TempFiles.Count -gt 0) {
            $TempFiles | Remove-Item -Force
            Write-Host "Removed $($TempFiles.Count) temporary files"
        }
        
        return $true
    }
    catch {
        Write-Warning "Cleanup operation failed: $($_.Exception.Message)"
        return $false
    }
}

# Export functions for use in main script
Export-ModuleMember -Function @(
    'Test-BulkLitigationHoldPrerequisites',
    'Get-OptimalBatchConfiguration', 
    'Start-PerformanceMonitoring',
    'Stop-PerformanceMonitoring',
    'Test-LitigationHoldLicense',
    'Export-BulkOperationReport',
    'Invoke-BulkLitigationHoldCleanup'
)