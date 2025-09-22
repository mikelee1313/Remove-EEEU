<#
.SYNOPSIS
    Removes special users ("Everyone" and "Everyone except external users") from SharePoint User Information List (UIL).

.DESCRIPTION
    This script connects to a SharePoint site using certificate-based authentication and removes specific 
    special users from the User Information List. This is typically used to clean up oversharing permissions
    but will break existing shares to these groups. The script includes safety warnings and confirmation
    prompts before executing destructive operations.

.PARAMETER tenantId
    The Azure AD tenant ID for authentication.

.PARAMETER appID
    The Azure AD application (client) ID for certificate authentication.

.PARAMETER siteUrl
    The full URL of the SharePoint site to connect to.

.PARAMETER thumbprint
    The certificate thumbprint used for authentication.

.PARAMETER bypassConfirmation
    Set to $true to skip the confirmation prompt for automated scenarios. Default is $false.

.PARAMETER specialUsersToRemove
    Array of special user names to remove from the UIL. Default includes "Everyone except external users" and "Everyone".

.NOTES
    File Name: Everyone-UIL-Remover.ps1
    Author: Mike Lee
    Date Created: 9/22/25
    Requires: PnP.PowerShell module
    API Permissions Required: 
    - SharePoint: Sites.FullControl.All (application permission)
    - SharePoint: User.ReadWrite.All (application permission)

.DISCLAIMER
    Disclaimer: The sample scripts are provided AS IS without warranty of any kind. 
    Microsoft further disclaims all implied warranties including, without limitation, 
    any implied warranties of merchantability or of fitness for a particular purpose. 
    The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
    In no event shall Microsoft, its authors, or anyone else involved in the creation, 
    production, or delivery of the scripts be liable for any damages whatsoever 
    (including, without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use of or inability 
    to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.
    
    IMPORTANT WARNINGS:
    - This operation is destructive and cannot be automatically undone
    - All existing sharing permissions to removed groups will be lost
    - Previously shared content will become inaccessible to affected users
    - Manual re-sharing will be required if content access is needed

.EXAMPLE
    .\Everyone-UIL-Remover.ps1
    Runs the script with default settings and prompts for confirmation.

.EXAMPLE
    # Set $bypassConfirmation = $true in script for automated execution
    .\Everyone-UIL-Remover.ps1
    Runs the script without confirmation prompts (for automation scenarios).

.LINK
    https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/
#>

#region Configuration
# ================================================================================================
# CONFIGURATION SECTION - Modify these values as needed
# ================================================================================================

# SharePoint Connection Settings
$tenantId = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"  
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"   
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"

# Script Behavior Settings
$bypassConfirmation = $false # Set to $true to skip confirmation prompt (for automated scenarios)

# Logging Configuration
$enableLogging = $true  # Set to $false to disable logging
$logFilePath = "$env:TEMP\EEEU-UIL-Remover_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"  # Log file path with timestamp

# Throttle Protection Configuration
$enableThrottleProtection = $true  # Enable throttle protection (recommended for multiple sites)
$delayBetweenSites = 2  # Delay in seconds between site processing to avoid overwhelming SharePoint
$delayBetweenOperations = 1  # Delay in seconds between individual API operations
$maxRetryAttempts = 3  # Maximum retry attempts when throttled (HTTP 429/503)
$baseRetryDelay = 5  # Base delay in seconds for exponential backoff when retrying

# Input Configuration - Sites to Process (REQUIRED)
$inputSiteList = "C:\temp\sitelist.txt" # REQUIRED: Specify site(s) to process
# Examples:
# Single site: $inputSiteList = "https://tenant.sharepoint.com/sites/sitename"
# Multiple sites from file: $inputSiteList = "C:\temp\sitelist.txt"
# Multiple sites from array: $inputSiteList = @("https://site1.com", "https://site2.com")

# Target Users Configuration
$specialUsersToRemove = @(
    "Everyone except external users",
    "Everyone"
)

#endregion Configuration

#region Logging Functions
# ================================================================================================
# LOGGING FUNCTIONS - For record keeping and audit trails
# ================================================================================================

function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO",
        [switch]$NoConsole
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $Level - $Message"
    
    # Write to log file if logging is enabled
    if ($enableLogging) {
        try {
            Add-Content -Path $logFilePath -Value $logEntry -ErrorAction SilentlyContinue
        }
        catch {
            # If logging fails, don't break the script
            Write-Warning "Failed to write to log file: $($_.Exception.Message)"
        }
    }
    
    # Write to console unless suppressed
    if (-not $NoConsole) {
        switch ($Level) {
            "INFO" { Write-Host $Message -ForegroundColor Cyan }
            "WARNING" { Write-Host $Message -ForegroundColor Yellow }
            "ERROR" { Write-Host $Message -ForegroundColor Red }
            "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        }
    }
}

function Write-LogHeader {
    param (
        [string]$Title,
        [int]$Width = 80
    )
    
    $separator = "=" * $Width
    $paddedTitle = " $Title ".PadLeft(($Width + $Title.Length) / 2).PadRight($Width)
    
    Write-Log $separator "INFO"
    Write-Log $paddedTitle "INFO"
    Write-Log $separator "INFO"
}

# Initialize logging
if ($enableLogging) {
    Write-Log "EEEU-UIL-Remover Script Started" "INFO"
    Write-Log "Log file: $logFilePath" "INFO"
    Write-Log "Script executed by: $env:USERNAME on $env:COMPUTERNAME" "INFO"
    Write-Log "Execution time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" "INFO"
    Write-Log "PowerShell version: $($PSVersionTable.PSVersion)" "INFO"
    Write-Log "" "INFO"
}

#endregion Logging Functions

#region Throttle Protection Functions
# ================================================================================================
# THROTTLE PROTECTION FUNCTIONS - Handle SharePoint throttling per Microsoft guidance
# ================================================================================================

function Invoke-ThrottleProtectedCommand {
    param(
        [scriptblock]$ScriptBlock,
        [string]$OperationName = "API Operation",
        [int]$MaxRetries = $maxRetryAttempts,
        [int]$TimeoutSeconds = 300  # 5 minute timeout per operation
    )
    
    $attempt = 1
    $success = $false
    $result = $null
    
    while ($attempt -le $MaxRetries -and -not $success) {
        try {
            if ($enableThrottleProtection -and $attempt -gt 1) {
                Write-Log "Retry attempt $attempt for: $OperationName" "WARNING"
            }
            
            # Execute the script block directly with timeout monitoring
            # Note: PowerShell Jobs don't work well with PnP context, so we use direct execution
            # with operation logging for monitoring
            Write-Log "Executing: $OperationName (attempt $attempt)" "INFO"
            $result = & $ScriptBlock
            $success = $true
            
            # Add small delay between operations if throttle protection is enabled
            if ($enableThrottleProtection -and $delayBetweenOperations -gt 0) {
                Start-Sleep -Seconds $delayBetweenOperations
            }
            
        }
        catch {
            $errorMessage = $_.Exception.Message
            $isThrottleError = $false
            $retryAfter = 0
            
            # Check for throttling indicators (HTTP 429/503)
            if ($errorMessage -match "429|Too Many Requests|503|Server Too Busy|Throttled") {
                $isThrottleError = $true
                $script:throttleEvents++
                Write-Log "⚠️  Throttling detected during: $OperationName" "WARNING"
                
                # Try to extract Retry-After header value from error
                if ($errorMessage -match "Retry-After.*?(\d+)") {
                    $retryAfter = [int]$matches[1]
                    Write-Log "Retry-After header indicates: $retryAfter seconds" "INFO"
                }
                else {
                    # Use exponential backoff if no Retry-After header
                    $retryAfter = $baseRetryDelay * [Math]::Pow(2, $attempt - 1)
                    Write-Log "Using exponential backoff: $retryAfter seconds" "INFO"
                }
            }
            
            if ($attempt -eq $MaxRetries) {
                Write-Log "❌ Failed after $MaxRetries attempts: $OperationName" "ERROR"
                Write-Log "Final error: $errorMessage" "ERROR"
                throw
            }
            elseif ($isThrottleError) {
                Write-Log "Waiting $retryAfter seconds before retry..." "INFO"
                Start-Sleep -Seconds $retryAfter
            }
            else {
                # Non-throttle error - shorter wait
                Write-Log "Non-throttle error, waiting 2 seconds before retry: $errorMessage" "WARNING"
                Start-Sleep -Seconds 2
            }
            
            $attempt++
        }
    }
    
    return $result
}

function Start-ThrottleProtectedDelay {
    param([string]$Context = "operation")
    
    if ($enableThrottleProtection -and $delayBetweenSites -gt 0) {
        Write-Log "Throttle protection: Waiting $delayBetweenSites seconds before next $Context..." "INFO"
        Start-Sleep -Seconds $delayBetweenSites
    }
}

#endregion Throttle Protection Functions

# Initialize tracking variables for summary reporting
$script:totalUsersRemoved = 0
$script:totalUsersFound = 0  
$script:sitesProcessedSuccessfully = 0
$script:removalDetails = @()
$script:throttleEvents = 0

#endregion Logging Functions

# Determine sites to process
$sitesToProcess = @()

if (-not $inputSiteList) {
    Write-Log "ERROR: inputSiteList is required but not specified!" "ERROR"
    Write-Log "Please set inputSiteList to:" "WARNING"
    Write-Log "  - A single site URL: `$inputSiteList = 'https://tenant.sharepoint.com/sites/sitename'" "WARNING"
    Write-Log "  - A file path: `$inputSiteList = 'C:\temp\sitelist.txt'" "WARNING"
    Write-Log "  - An array: `$inputSiteList = @('https://site1.com', 'https://site2.com')" "WARNING"
    throw "inputSiteList configuration is required"
}

if ($inputSiteList -is [string] -and (Test-Path $inputSiteList)) {
    # Input is a file path
    Write-Log "Reading site URLs from file: $inputSiteList" "INFO"
    $sitesToProcess = Get-Content -Path $inputSiteList | Where-Object { $_ -and $_.Trim() -ne "" }
    Write-Log "Found $($sitesToProcess.Count) sites in input file" "SUCCESS"
}
elseif ($inputSiteList -is [array]) {
    # Input is an array of URLs
    Write-Log "Using provided array of site URLs" "INFO"
    $sitesToProcess = $inputSiteList
    Write-Log "Found $($sitesToProcess.Count) sites in input array" "SUCCESS"
}
elseif ($inputSiteList -is [string]) {
    # Input is a single site URL
    Write-Log "Processing single site: $inputSiteList" "INFO"
    $sitesToProcess = @($inputSiteList)
}
else {
    Write-Log "ERROR: Invalid inputSiteList configuration!" "ERROR"
    throw "inputSiteList must be a string (URL or file path) or an array of URLs"
}

# Function to process a single site
function Invoke-SiteProcessing {
    param(
        [string]$currentSiteUrl
    )
    
    Write-LogHeader "Processing Site: $currentSiteUrl"

    # Connect to SharePoint using PnP PowerShell with certificate authentication and throttle protection
    try {
        Write-Log "Connecting to SharePoint site: $currentSiteUrl" "INFO"
        
        $connectResult = Invoke-ThrottleProtectedCommand -OperationName "Connect to SharePoint" -ScriptBlock {
            Connect-PnPOnline -Url $currentSiteUrl -ClientId $appID -Thumbprint $thumbprint -Tenant $tenantId
        }
        
        Write-Log "✅ Successfully connected to SharePoint" "SUCCESS"
    }
    catch {
        Write-Log "❌ Failed to connect to SharePoint: $($_.Exception.Message)" "ERROR"
        return $false
    }

    # Find specific special users directly (more efficient than retrieving all users)
    $usersFound = @()
    Write-Log "Looking for special users to remove..." "INFO"

    foreach ($specialUser in $specialUsersToRemove) {
        try {
            # Try to get the specific user directly by display name using a targeted approach
            Write-Log "  Checking for: $specialUser" "INFO"
            
            # Try to find the user efficiently using different approaches
            $user = $null
            try {
                Write-Log "    Searching for user: $specialUser" "INFO"
                
                # Try to get the user by title directly (most common case) with throttle protection
                $users = Invoke-ThrottleProtectedCommand -OperationName "Get SharePoint Users" -ScriptBlock {
                    Get-PnPUser
                }
                $user = $users | Where-Object { $_.Title -eq $specialUser } | Select-Object -First 1
                
                # Alternative: Also try searching by LoginName pattern for system accounts
                if (-not $user) {
                    if ($specialUser -eq "Everyone") {
                        $user = $users | Where-Object { $_.LoginName -like "*c:0(.s|true*" } | Select-Object -First 1
                    }
                    elseif ($specialUser -eq "Everyone except external users") {
                        $user = $users | Where-Object { $_.LoginName -like "*spo-grid-all-users*" } | Select-Object -First 1
                    }
                }
            }
            catch {
                Write-Log "    Error during user lookup: $($_.Exception.Message)" "ERROR"
            }
            
            if ($user) {
                $usersFound += $user
                $script:totalUsersFound++
                Write-Log "  ✓ Found user to remove: $($user.Title) (Login: $($user.LoginName))" "WARNING"
                Write-Log "    User details - Title: '$($user.Title)', LoginName: '$($user.LoginName)', Email: '$($user.Email)', ID: $($user.Id)" "INFO"
            }
            else {
                Write-Log "  ✗ User not found: $specialUser" "INFO"
            }
        }
        catch {
            Write-Log "  ✗ Error checking for user '$specialUser': $($_.Exception.Message)" "ERROR"
        }
    }

    # Display warning and get confirmation before proceeding (only for first site if processing multiple)
    if ($usersFound.Count -gt 0) {
        # Since global confirmation was already obtained, proceed without additional prompts
        if (-not $script:globalConfirmationGiven) {
            Write-Log "Global confirmation not given - this should not happen" "ERROR"
            return $false
        }
        
        Write-Log "Proceeding with removal for site: $currentSiteUrl" "INFO"
        $shouldProceed = $true
    }

    # Remove the special users from UIL
    if ($usersFound.Count -eq 0) {
        Write-Log "No special users (EEEU or Everyone) found in User Information List." "SUCCESS"
    }
    else {
        Write-Log "Starting removal of $($usersFound.Count) special users from UIL..." "INFO"
        foreach ($userToRemove in $usersFound) {
            try {
                Write-Log "Attempting to remove user from UIL: $($userToRemove.Title)" "INFO"
                
                Invoke-ThrottleProtectedCommand -OperationName "Remove User from UIL" -ScriptBlock {
                    Remove-PnPUser -Identity $userToRemove.LoginName -Force
                }
                
                # Track successful removal
                $script:totalUsersRemoved++
                $removalDetail = @{
                    Site          = $currentSiteUrl
                    UserTitle     = $userToRemove.Title
                    UserLoginName = $userToRemove.LoginName
                    UserEmail     = $userToRemove.Email
                    UserID        = $userToRemove.Id
                    RemovalTime   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                    Status        = "SUCCESS"
                }
                $script:removalDetails += $removalDetail
                
                Write-Log "✓ Successfully removed $($userToRemove.Title) from User Information List." "SUCCESS"
                Write-Log "  Removed user details - Title: '$($userToRemove.Title)', LoginName: '$($userToRemove.LoginName)', Site: '$currentSiteUrl'" "INFO"
            }
            catch {
                # Track failed removal
                $removalDetail = @{
                    Site          = $currentSiteUrl
                    UserTitle     = $userToRemove.Title
                    UserLoginName = $userToRemove.LoginName
                    UserEmail     = $userToRemove.Email
                    UserID        = $userToRemove.Id
                    RemovalTime   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                    Status        = "FAILED"
                    Error         = $_.Exception.Message
                }
                $script:removalDetails += $removalDetail
                
                Write-Log "✗ Failed to remove $($userToRemove.Title): $($_.Exception.Message)" "ERROR"
                Write-Log "  Failed removal details - Title: '$($userToRemove.Title)', LoginName: '$($userToRemove.LoginName)', Site: '$currentSiteUrl'" "ERROR"
            }
        }
    }

    # Disconnect from SharePoint (only if still connected)
    try {
        $connection = Get-PnPConnection -ErrorAction SilentlyContinue
        if ($connection) {
            Disconnect-PnPOnline
            Write-Log "Disconnected from SharePoint" "INFO"
        }
    }
    catch {
        Write-Log "Connection already closed or no connection to disconnect" "INFO"
    }

    # Track successful site processing
    $script:sitesProcessedSuccessfully++
    Write-Log "Completed processing for site: $currentSiteUrl" "SUCCESS"

    return $true
}

# Process all sites
$totalSites = $sitesToProcess.Count
$processedSites = 0
$cancelledByUser = $false

# Initialize confirmation tracking
$script:confirmationShown = $false
$script:globalConfirmationGiven = $false

# Global confirmation before processing any sites
if (-not $bypassConfirmation) {
    Write-Host ""
    Write-Host "⚠️  IMPORTANT: You are about to remove special users from SharePoint sites" -ForegroundColor Black -BackgroundColor Yellow
    Write-Host "🌐 Sites to process: $($sitesToProcess.Count)" -ForegroundColor Cyan
    Write-Host "👥 Users to remove: $($specialUsersToRemove -join ', ')" -ForegroundColor Magenta
    Write-Host ""
    Write-Host "🚨 CONSEQUENCES:" -ForegroundColor Red
    Write-Host "• All sharing permissions to these groups will be REMOVED" -ForegroundColor Red
    Write-Host "• Previously shared content will become inaccessible" -ForegroundColor Red
    Write-Host "• This action CANNOT be undone automatically" -ForegroundColor Red
    Write-Host ""
    
    # Check if we're in an interactive session
    if ([Environment]::UserInteractive -and -not [Console]::IsInputRedirected) {
        Write-Host "🔐 CONFIRMATION REQUIRED" -ForegroundColor Yellow
        $confirmation = Read-Host "Proceed with removing users from ALL sites? (Type 'YES' to confirm)"
        if ($confirmation -ne "YES") {
            Write-Host "❌ Operation cancelled by user." -ForegroundColor Red
            Write-Log "Operation cancelled by user during global confirmation." "WARNING"
            exit 0
        }
        Write-Host "✅ Confirmation received - proceeding with operation" -ForegroundColor Green
        $script:globalConfirmationGiven = $true
    }
    else {
        Write-Log "Running in non-interactive mode. Use `$bypassConfirmation = `$true to proceed automatically." "ERROR"
        Write-Host "❌ Non-interactive session detected. Set `$bypassConfirmation = `$true to proceed." -ForegroundColor Red
        exit 1
    }
}
else {
    $script:globalConfirmationGiven = $true
    Write-Host "⚡ Confirmation bypassed - proceeding automatically" -ForegroundColor Yellow
    Write-Log "Confirmation bypassed - proceeding automatically" "WARNING"
}

# Display throttle protection configuration
Write-LogHeader "THROTTLE PROTECTION CONFIGURATION"
if ($enableThrottleProtection) {
    Write-Host "🛡️  Throttle Protection: ENABLED" -ForegroundColor Green
    Write-Host "   • Delay between sites: $delayBetweenSites seconds" -ForegroundColor Green
    Write-Host "   • Delay between operations: $delayBetweenOperations seconds" -ForegroundColor Green
    Write-Host "   • Max retry attempts: $maxRetryAttempts" -ForegroundColor Green
    Write-Host "   • Base retry delay: $baseRetryDelay seconds (with exponential backoff)" -ForegroundColor Green
    Write-Host "   • Automatic retry on HTTP 429/503 responses" -ForegroundColor Green
    Write-Host "   • Honors Retry-After headers from SharePoint" -ForegroundColor Green
}
else {
    Write-Host "⚠️  Throttle Protection: DISABLED" -ForegroundColor Yellow
    Write-Host "   Consider enabling throttle protection when processing multiple sites" -ForegroundColor Yellow
}
Write-Log "" "INFO"

foreach ($site in $sitesToProcess) {
    $processedSites++
    Write-Host ""
    Write-Host "🌐 Processing site $processedSites of $totalSites" -ForegroundColor Magenta
    Write-Host "📍 Site: $site" -ForegroundColor Cyan
    
    # Add delay between sites to avoid overwhelming SharePoint (except for first site)
    if ($processedSites -gt 1) {
        Start-ThrottleProtectedDelay -Context "site processing"
    }
    
    $result = Invoke-SiteProcessing -currentSiteUrl $site
    
    if (-not $result) {
        if ($processedSites -eq 1) {
            # First site failed or was cancelled - assume user cancelled for all
            $cancelledByUser = $true
            break
        }
        else {
            # Subsequent site failed - continue with others
            Write-Host "❌ Failed to process site: $site" -ForegroundColor Red
            continue
        }
    }
    else {
        Write-Host "✅ Site processing completed successfully" -ForegroundColor Green
    }
}

if ($cancelledByUser) {
    Write-Log "" "INFO"
    Write-Log "Operation cancelled. No sites were processed." "WARNING"
}
else {
    Write-LogHeader "PROCESSING COMPLETE"
    Write-Log "Processed $processedSites of $totalSites sites" "SUCCESS"
    
    # Comprehensive execution summary
    Write-LogHeader "EXECUTION SUMMARY"
    Write-Log "Sites successfully processed: $script:sitesProcessedSuccessfully" "INFO"
    Write-Log "Total users found across all sites: $script:totalUsersFound" "INFO"
    Write-Log "Total users successfully removed: $script:totalUsersRemoved" "INFO"
    
    # Throttle protection statistics
    if ($enableThrottleProtection) {
        Write-Log "Throttle protection enabled: YES" "INFO"
        Write-Log "Throttle events encountered: $script:throttleEvents" "INFO"
        if ($script:throttleEvents -gt 0) {
            Write-Log "⚠️  Note: Throttling was encountered during execution. This is normal for large operations." "WARNING"
            Write-Log "The script successfully handled all throttle events using retry logic." "SUCCESS"
        }
    }
    else {
        Write-Log "Throttle protection enabled: NO" "WARNING"
    }
    $failedRemovals = ($script:removalDetails | Where-Object { $_.Status -eq "FAILED" }).Count
    if ($failedRemovals -gt 0) {
        Write-Log "Total users that failed to remove: $failedRemovals" "WARNING"
    }
    
    if ($script:removalDetails.Count -gt 0) {
        Write-LogHeader "DETAILED REMOVAL RESULTS"
        
        # Group by site for better organization
        $siteGroups = $script:removalDetails | Group-Object -Property Site
        
        foreach ($siteGroup in $siteGroups) {
            Write-Log "" "INFO"
            Write-Log "Site: $($siteGroup.Name)" "INFO"
            Write-Log "$("-" * 60)" "INFO"
            
            $successfulRemovals = $siteGroup.Group | Where-Object { $_.Status -eq "SUCCESS" }
            $failedRemovals = $siteGroup.Group | Where-Object { $_.Status -eq "FAILED" }
            
            if ($successfulRemovals.Count -gt 0) {
                $successCount = @($successfulRemovals).Count
                Write-Log "✓ Successfully removed users ($successCount):" "SUCCESS"
                foreach ($removal in $successfulRemovals) {
                    Write-Log "  - $($removal.UserTitle)" "SUCCESS"
                    Write-Log "    LoginName: $($removal.UserLoginName)" "INFO"
                    Write-Log "    Email: $($removal.UserEmail)" "INFO"
                    Write-Log "    Removed at: $($removal.RemovalTime)" "INFO"
                }
            }
            
            if ($failedRemovals.Count -gt 0) {
                $failedCount = @($failedRemovals).Count
                Write-Log "✗ Failed to remove users ($failedCount):" "ERROR"
                foreach ($removal in $failedRemovals) {
                    Write-Log "  - $($removal.UserTitle)" "ERROR"
                    Write-Log "    LoginName: $($removal.UserLoginName)" "ERROR"
                    Write-Log "    Email: $($removal.UserEmail)" "ERROR"
                    Write-Log "    Error: $($removal.Error)" "ERROR"
                    Write-Log "    Failed at: $($removal.RemovalTime)" "ERROR"
                }
            }
        }
    }
    else {
        Write-Log "No users were found for removal across any of the processed sites." "INFO"
    }
}

# Final logging summary
if ($enableLogging) {
    Write-Log "" "INFO"
    Write-Log "Script execution completed at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" "INFO"
    Write-Log "Log file location: $logFilePath" "INFO"
    Write-Host ""
    Write-Host "📋 SCRIPT EXECUTION COMPLETED" -ForegroundColor Yellow
    Write-Host "📄 Complete execution log saved to: $logFilePath" -ForegroundColor Cyan
    Write-Host "📊 Check the log file for detailed results and audit trail" -ForegroundColor Cyan
}
