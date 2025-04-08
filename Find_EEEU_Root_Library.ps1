<#
.SYNOPSIS
Scans OneDrive sites for Everyone Except External Users (EEEU) permissions at root and document library level.

.DESCRIPTION
This script connects to SharePoint Online and scans all OneDrive sites for permissions granted to 
"Everyone Except External Users" (identified by the claim "spo-grid-all-users"). The script checks 
both at the root site level and the default document library level.

The script requires the PnP.PowerShell module and an app registration with Sites.FullControl.All permissions.

.PARAMETER None
This script does not accept pipeline input or parameters. Configuration is hard-coded in the script.

.OUTPUTS
The script generates two outputs:
1. A log file at $env:TEMP\Find_EEEU_Root_Library_[timestamp].txt
2. A CSV report at $env:TEMP\OneDrive_EEEU_Root_Library_[timestamp].csv with details of any EEEU permissions found

.NOTES
File Name      : Find_EEEU_Root_Library.ps1
Author         : Mike Lee
Date Created   : 3/27/25
Date Updated   : 4/8/25 - Added Suport to handle throttling.
Prerequisite   : PnP.PowerShell module
                 App registration with Sites.FullControl.All permission

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

.EXAMPLE
PS> .\Find_EEEU_Root_Library.ps1
Connects to SharePoint Online and scans all OneDrive sites for EEEU permissions.

.FUNCTIONALITY
- Connects to SharePoint Online admin center using app-only authentication
- Retrieves all OneDrive sites in the tenant
- For each OneDrive site, checks if EEEU permissions exist at the root web level
- For each OneDrive site, checks if EEEU permissions exist on the default document library
- Logs all activities and results to a log file
- Exports findings to a CSV file
- Handles API throttling with automatic retries based on Retry-After headers
#>

$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"
$spadmin = "m365cpi13246019"

$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "$env:TEMP\Find_EEEU_Root_Library_$startime.txt"
$maxRetries = 5  # Maximum number of retries for throttled requests

function Write-Log {
    param (
        [string]$message,
        [string]$level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $level - $message"
    Add-Content -Path $logFilePath -Value $logMessage
    
    # Also output to console based on log level
    switch ($level) {
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "INFO" { Write-Host $logMessage -ForegroundColor White }
        "THROTTLE" { Write-Host $logMessage -ForegroundColor Cyan }
        default { Write-Host $logMessage }
    }
}

function Invoke-PnPWithRetry {
    param (
        [scriptblock]$ScriptBlock,
        [string]$Operation,
        [int]$MaxRetries = $maxRetries
    )
    
    $retryCount = 0
    $success = $false
    $result = $null
    
    while (-not $success -and $retryCount -lt $MaxRetries) {
        try {
            $result = & $ScriptBlock
            $success = $true
        }
        catch {
            # Check if the error is due to throttling (429 Too Many Requests)
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode -eq 429) {
                $retryCount++
                
                # Extract retry-after header if present
                $retryAfterSeconds = 10  # Default retry delay
                if ($_.Exception.Response.Headers -and $_.Exception.Response.Headers["Retry-After"]) {
                    $retryAfterSeconds = [int]$_.Exception.Response.Headers["Retry-After"]
                }
                
                Write-Log "Throttling detected during $Operation. Retry $retryCount of $MaxRetries. Waiting $retryAfterSeconds seconds..." -level "THROTTLE"
                Start-Sleep -Seconds $retryAfterSeconds
            }
            else {
                # If it's not a throttling error, rethrow
                Write-Log "Error during $Operation : $($_.Exception.Message)" -level "ERROR"
                throw $_
            }
        }
    }
    
    if (-not $success) {
        Write-Log "Failed to complete $Operation after $MaxRetries retries" -level "ERROR"
        throw "Failed to complete $Operation after $MaxRetries retries"
    }
    
    return $result
}

# Connect to SharePoint Admin Center with retry logic
Write-Log "Connecting to SharePoint Admin Center" -level "INFO"
Invoke-PnPWithRetry -ScriptBlock {
    Connect-PnPOnline -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant -Url "https://$spadmin-admin.sharepoint.com" -ErrorAction Stop
} -Operation "Connect to SharePoint Admin Center"

# Retrieve OneDrive sites with retry logic
Write-Log "Retrieving OneDrive sites" -level "INFO"
$personalSites = Invoke-PnPWithRetry -ScriptBlock {
    Get-PnPTenantSite -IncludeOneDriveSites | Where-Object -Property Template -match "SPSPERS"
} -Operation "Retrieve OneDrive sites"

# Function to check for EEEU at the default document library level
function Test-EEEUDocumentLibrary {
    param (
        [string]$siteUrl,
        [object]$personalSite
    )

    #Write-Log "Checking for EEEU: $siteUrl" -level "INFO"

    try {
        # Connect to the site with retry logic
        Invoke-PnPWithRetry -ScriptBlock {
            Connect-PnPOnline -Url $siteUrl -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant -ErrorAction Stop
        } -Operation "Connect to site $siteUrl"

        # Get the default document library with retry logic
        $documentLibrary = Invoke-PnPWithRetry -ScriptBlock {
            Get-PnPList -Identity "documents"
        } -Operation "Get document library at $siteUrl"

        # Get RoleAssignments for the document library with retry logic
        $roleAssignments = Invoke-PnPWithRetry -ScriptBlock {
            Get-PnPProperty -ClientObject $documentLibrary -Property RoleAssignments
        } -Operation "Get role assignments for document library at $siteUrl"

        # Enumerate RoleAssignments
        foreach ($roleAssignment in $roleAssignments) {
            # Pull principal LoginName with retry logic
            Invoke-PnPWithRetry -ScriptBlock {
                Get-PnPProperty -ClientObject $roleAssignment.Member -Property LoginName | Out-Null
            } -Operation "Get member login name at $siteUrl"

            # Enumerate principals
            foreach ($member in $roleAssignment.Member) {
                # Check if principal is EEEU
                if ($member.LoginName -match "spo-grid-all-users") {
                    # Hydrate RoleDefinitionBindings (permissions) with retry logic
                    Invoke-PnPWithRetry -ScriptBlock {
                        Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings | Out-Null
                    } -Operation "Get role definition bindings at $siteUrl"

                    # Filter out hidden permissions and Limited Access
                    $roleDefinitionBindings = $roleAssignment.RoleDefinitionBindings | Where-Object { 
                        -not $_.Hidden -and $_.Name -ne "Limited Access" 
                    }
                    
                    if ($roleDefinitionBindings) {
                        Write-Log "Found EEEU in Document Library: $($member.LoginName) on $siteUrl" -level "WARNING"
                        
                        # Output object
                        [PSCustomObject] @{
                            SiteUrl                = $siteUrl
                            Owner                  = $personalSite.Owner
                            Claim                  = $member.LoginName
                            RoleDefinitionBindings = $roleDefinitionBindings.Name -join ","
                            LibraryName            = $documentLibrary.Title
                        }
                    }
                    else { 
                        continue 
                    }
                }
            }
        }
    }
    catch {
        Write-Log "Error processing document library: $($_)" -level "ERROR"
    }
}

# Initialize results array
$results = @()
$processedCount = 0
$totalSites = $personalSites.Count

# Enumerate OneDrive sites
foreach ($personalSite in $personalSites) {
    $processedCount++
    $percentComplete = [math]::Round(($processedCount / $totalSites) * 100, 2)
    Write-Log "Processing site $processedCount of $totalSites ($percentComplete%): $($personalSite.Url)" -level "INFO"
    
    try {
        # Connect to the OneDrive site with retry logic
        Invoke-PnPWithRetry -ScriptBlock {
            Connect-PnPOnline -Url $personalSite.Url -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant -ErrorAction Stop
        } -Operation "Connect to site $($personalSite.Url)"

        # Get the rootweb + RoleAssignments with retry logic
        $web = Invoke-PnPWithRetry -ScriptBlock {
            Get-PnPWeb -Includes RoleAssignments
        } -Operation "Get web for site $($personalSite.Url)"

        # Enumerate RoleAssignments
        foreach ($roleAssignment in $web.RoleAssignments) {
            # Pull principal LoginName with retry logic
            Invoke-PnPWithRetry -ScriptBlock {
                Get-PnPProperty -ClientObject $roleAssignment.Member -Property LoginName |  Out-Null
            } -Operation "Get member login name for site $($personalSite.Url)"

            # Enumerate principals
            foreach ($member in $roleAssignment.Member) {
                # Check if principal is EEEU
                if ($member.LoginName -match "spo-grid-all-users") {
                    # Hydrate RoleDefinitionBindings (permissions) with retry logic
                    Invoke-PnPWithRetry -ScriptBlock {
                        Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings | Out-Null
                    } -Operation "Get role definition bindings for site $($personalSite.Url)"

                    # Filter out hidden permissions and Limited Access
                    $roleDefinitionBindings = $roleAssignment.RoleDefinitionBindings | Where-Object { 
                        -not $_.Hidden -and $_.Name -ne "Limited Access" 
                    }
                    
                    if ($roleDefinitionBindings) {
                        Write-Log "Found EEEU: $($member.LoginName) on $($personalSite.Url)" -level "WARNING"
                        
                        # Create and add result object
                        $result = [PSCustomObject] @{
                            SiteUrl                = $personalSite.Url
                            Owner                  = $personalSite.Owner
                            Claim                  = $member.LoginName
                            RoleDefinitionBindings = $roleDefinitionBindings.Name -join ","
                            LibraryName            = "RootWeb"
                        }
                        $results += $result
                    }
                    else { 
                        continue 
                    }
                }
            }
        }

        # Call the function to check for EEEU at the default document library level
        $libraryResults = Test-EEEUDocumentLibrary -siteUrl $personalSite.Url -personalSite $personalSite
        if ($libraryResults) {
            $results += $libraryResults
        }
    }
    catch {
        Write-Log "Error processing site $($personalSite.Url): $($_)" -level "ERROR"
    }
}

# Export results to CSV
$csvPath = "$env:TEMP\OneDrive_EEEU_Root_Library_$startime.csv"
$results | Export-Csv -Path $csvPath -NoTypeInformation
Write-Log "Scan completed. Processed $processedCount sites. Results exported to: $csvPath" -level "INFO"
Write-Log "Log file saved to: $logFilePath" -level "INFO"
