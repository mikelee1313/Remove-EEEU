<#
.SYNOPSIS
    Scans SharePoint Online sites to identify occurrences of the "Everyone Except External Users" (EEEU) group in file permissions.

.DESCRIPTION
    This script connects to SharePoint Online using provided tenant-level credentials and iterates through a list of 
    site URLs specified in an input file. It recursively scans document libraries and lists (excluding specified folders) 
    to locate files where the "Everyone Except External Users" group has permissions assigned (excluding "Limited Access"). 
    The script logs its operations and outputs the results to a CSV file, detailing the site URL, file URL, and assigned roles.

.PARAMETER None
    This script does not accept parameters via the command line. Configuration is done within the script.

.INPUTS
    A text file containing SharePoint site URLs to scan (path specified in $inputFilePath variable).

.OUTPUTS
    - A CSV file containing all found EEEU occurrences (path: $env:TEMP\Find_EEEU_In_Sites_[timestamp].csv)
    - A log file documenting the script's execution (path: $env:TEMP\Find_EEEU_In_Sites_[timestamp].txt)

.NOTES
    File Name      : Find-EEEUInSites.ps1
    Author         : Mike Lee
    Date Created   : 5/12/2025

    The script uses app-only authentication with a certificate thumbprint. Make sure the app has
    proper permissions in your tenant (Sites.FullControl.All is recommended).

    The script ignores several system folders and lists to improve performance and avoid errors.

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
    .\Find-EEEUInSites.ps1
    Executes the script with the configured settings. Ensure you've updated the variables at the top
    of the script (appID, thumbprint, tenant, and inputFilePath) before running.
#>
# Tenant Level Information
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"

# Script Parameters
Add-Type -AssemblyName System.Web
$EEEU = '*spo-grid-all-users*'
$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "$env:TEMP\Find_EEEU_In_Sites_$startime.txt"
$outputFilePath = "$env:TEMP\Find_EEEU_In_Sites_$startime.csv"

# Path and file names
$inputFilePath = "C:\temp\oversharedurls.txt" # Path to the input file containing site URLs

# Setup logging
function Write-Log {
    param (
        [string]$message,
        [string]$level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $level - $message"
    Add-Content -Path $logFilePath -Value $logMessage
}

# Handle SharePoint Online throttling with exponential backoff
function Invoke-WithRetry {
    param (
        [ScriptBlock]$ScriptBlock,
        [int]$MaxRetries = 5,
        [int]$InitialDelaySeconds = 5
    )
    
    $retryCount = 0
    $delay = $InitialDelaySeconds
    $success = $false
    $result = $null
    
    while (-not $success -and $retryCount -lt $MaxRetries) {
        try {
            $result = & $ScriptBlock
            $success = $true
        }
        catch {
            $exception = $_.Exception
            
            # Check if this is a throttling error (look for specific status codes or messages)
            $isThrottlingError = $false
            $retryAfterSeconds = $delay
            
            if ($exception.Response -ne $null) {
                # Check for Retry-After header
                $retryAfterHeader = $exception.Response.Headers['Retry-After']
                if ($retryAfterHeader) {
                    $isThrottlingError = $true
                    $retryAfterSeconds = [int]$retryAfterHeader
                    Write-Log "Received Retry-After header: $retryAfterSeconds seconds" "WARNING"
                }
                
                # Check for 429 (Too Many Requests) or 503 (Service Unavailable)
                $statusCode = [int]$exception.Response.StatusCode
                if ($statusCode -eq 429 -or $statusCode -eq 503) {
                    $isThrottlingError = $true
                    Write-Log "Detected throttling response (Status code: $statusCode)" "WARNING"
                }
            }
            
            # Also check for specific throttling error messages
            if ($exception.Message -match "throttl" -or 
                $exception.Message -match "too many requests" -or
                $exception.Message -match "temporarily unavailable") {
                $isThrottlingError = $true
                Write-Log "Detected throttling error in message: $($exception.Message)" "WARNING"
            }
            
            if ($isThrottlingError) {
                $retryCount++
                if ($retryCount -lt $MaxRetries) {
                    Write-Log "Throttling detected. Retry attempt $retryCount of $MaxRetries. Waiting $retryAfterSeconds seconds..." "WARNING"
                    Write-Host "Throttling detected. Retry attempt $retryCount of $MaxRetries. Waiting $retryAfterSeconds seconds..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $retryAfterSeconds
                    
                    # Implement exponential backoff if no Retry-After header was provided
                    if ($retryAfterSeconds -eq $delay) {
                        $delay = $delay * 2 # Exponential backoff
                    }
                }
                else {
                    Write-Log "Maximum retry attempts reached. Giving up on operation." "ERROR"
                    throw $_
                }
            }
            else {
                # Not a throttling error, rethrow
                Write-Log "General Error occurred During retrieval : $($_.Exception.Message)" "WARNING"
                throw $_
            }
        }
    }
    
    return $result
}

# Read site URLs from input file
function Read-SiteURLs {
    param (
        [string]$filePath
    )
    $urls = Get-Content -Path $filePath
    return $urls
}

# Connect to SharePoint Online
function Connect-SharePoint {
    param (
        [string]$siteURL
    )
    try {
        Invoke-WithRetry -ScriptBlock {
            Connect-PnPOnline -Url $siteURL -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
        }
        Write-Log "Connected to SharePoint Online at $siteURL"
        return $true # Connection successful
    }
    catch {
        Write-Log "Failed to connect to SharePoint Online at $siteURL : $($_.Exception.Message)" "ERROR"
        return $false # Connection failed
    }
}

# List of folders to ignore
$ignoreFolders = @(
    "_catalogs",
    "_cts",
    "_private",
    "_vti_pvt",
    "Reference 778a30bb4f074ae3bec315889ee34b88",
    "Sharing Links",
    "Social",
    "FavoriteLists-e0157a47-72e4-43c1-bfd0-ed9f7040e894",
    "User Information List",
    "Web Template Extensions",
    "SmartCache-8189C6B3-4081-4F62-9015-35FDB7FDF042",
    "SharePointHomeCacheList",
    "RecentLists-56BAEAB4-E7AD-4E59-B92B-9290D871F5C3",
    "PersonalCacheLibrary",
    "microsoft.ListSync.Endpoints",
    "Maintenance Log Library",
    "DO_NOT_DELETE_ENTERPRISE_USER_CONTAINER_ENUM_LIST_ee0de9c4-6398-408f-ac09-f0401edfb0bf",
    "appfiles",
    "(Reference, 778a30bb4f074ae3bec315889ee34b88)"
)

# Get all items in a folder and subfolders (using absolute URLs)
function Get-AllItemsInFolderAbs {
    param (
        [string]$siteURL,
        [string]$folderUrl
    )
    try {
        $allItems = @()
        if ($folderUrl) {
            # Check if folderUrl is not empty
            $items = Invoke-WithRetry -ScriptBlock {
                Get-PnPListItem -List $folderUrl -PageSize 500
            }
            $allItems += $items | Where-Object { $_["FileLeafRef"] -like "*.*" }

            $subFolders = Invoke-WithRetry -ScriptBlock {
                Get-PnPFolderItem -FolderSiteRelativeUrl $folderUrl -ItemType Folder
            }
            foreach ($folder in $subFolders) {
                if ($folder.Name -notin $ignoreFolders) {
                    $allItems += Get-AllItemsInFolderAbs -siteURL $siteURL -folderUrl $folder.ServerRelativeUrl
                }
                else {
                    #Write-Log "Ignoring folder: $($folder.DisplayName)"
                }
            }
        }
        return $allItems
    }
    catch {
        Write-Log "Failed to retrieve items from folder $folderUrl : $_" "ERROR"
        return @()
    }
}

# Process file
function Find-EEEUinFiles {
    param (
        $item
    )
    try {
        $file = @()
        $fileUrl = $item.FieldValues.FileRef
        
        # Check if the file URL contains any of the ignore folders
        foreach ($ignoreFolder in $ignoreFolders) {
            if ($fileUrl -like "*/$ignoreFolder/*" -or $fileUrl -like "*/$ignoreFolder") {
                #Write-Log "Ignoring file: $fileUrl because it's in ignored folder: $ignoreFolder"
                return # Add return statement here to skip processing the ignored file
            }
        }
       
        try {
            # Try direct approach first with throttling protection
            $file = Invoke-WithRetry -ScriptBlock {
                Get-PnPFile -Url $fileUrl -AsListItem -ErrorAction Stop
            }
        }
        catch {
            # If direct approach fails, try with URL encoding
            try {
                Write-Log "Initial file access failed, trying with URL encoding: $fileUrl" "WARNING"
                
                # Parse the URL into parts
                $urlParts = $fileUrl.Split('/')
                
                # Encode each part of the URL separately (except the protocol and domain)
                $encodedParts = @()
                $skipEncoding = $true
                foreach ($part in $urlParts) {
                    # Skip encoding for the protocol and domain parts
                    if ($skipEncoding -and ($part -eq "https:" -or $part -eq "" -or $part -like "*.sharepoint.com")) {
                        $encodedParts += $part
                    }
                    else {
                        $skipEncoding = $false
                        $encodedParts += [System.Web.HttpUtility]::UrlEncode($part)
                    }
                }
                
                # Rebuild the URL with encoded parts
                $encodedFileUrl = $encodedParts -join '/'
                
                # Try with encoded URL and throttling protection
                $file = Invoke-WithRetry -ScriptBlock {
                    Get-PnPFile -Url $encodedFileUrl -AsListItem
                }
                Write-Log "Successfully accessed file with encoded URL: $encodedFileUrl"
            }
            catch {
                Write-Log "Failed to access file even with URL encoding: $fileUrl - $_" "ERROR"
                return
            }
        }

        # Get permissions with throttling protection
        $Permissions = Invoke-WithRetry -ScriptBlock {
            Get-PnPProperty -ClientObject $file -Property RoleAssignments
        }

        if ($Permissions) {
            $roles = @()
            foreach ($RoleAssignment in $Permissions) {
                # Get role assignments with throttling protection
                Invoke-WithRetry -ScriptBlock {
                    Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
                }
                
                #Write-Log "Checking File: $($fileUrl) role assignment: $($RoleAssignment.Member.LoginName), Role: $($RoleAssignment.RoleDefinitionBindings.name)"

                if ($RoleAssignment.Member.LoginName -like $EEEU -and $RoleAssignment.RoleDefinitionBindings.name -ne 'Limited Access') {
                    $rolelevel = $RoleAssignment.RoleDefinitionBindings
                    foreach ($role in $rolelevel) {
                        $roles += $role.Name
                    }   
                }
            }
            if ($roles.Count -gt 0) {
                # Get file owner information
                $owner = "Unknown"
                $ownerEmail = "Unknown"
                $createdDate = "Unknown"
                
                try {
                    # Try to get file author/owner information using PnP methods
                    if ($file.FieldValues.ContainsKey("Author")) {
                        $authorId = $file.FieldValues.Author.LookupId
                        
                        if ($authorId) {
                            $ownerInfo = Invoke-WithRetry -ScriptBlock {
                                Get-PnPUser -Identity $authorId
                            }
                            
                            if ($ownerInfo) {
                                $owner = $ownerInfo.Title
                                $ownerEmail = $ownerInfo.Email
                            }
                        }
                    }
                    
                    # Get created date
                    if ($file.FieldValues.ContainsKey("Created")) {
                        $createdDate = $file.FieldValues.Created.ToString("yyyy-MM-dd HH:mm:ss")
                    }
                }
                catch {
                    Write-Log "Error retrieving file owner information: $_" "WARNING"
                }

                $global:EEEUOccurrences += [PSCustomObject]@{
                    Url         = $SiteURL
                    ItemURL     = $file.FieldValues.FileRef
                    RoleNames   = ($roles -join ", ")
                    OwnerName   = $owner
                    OwnerEmail  = $ownerEmail
                    CreatedDate = $createdDate
                }
                Write-Host "Located EEEU in file: $($file.FieldValues.FileLeafRef) on $SiteURL" -ForegroundColor Red
                Write-Log "Located EEEU in file: $($file.FieldValues.FileLeafRef) on $SiteURL"
            }
        }
    }
    catch {
        Write-Log "Failed to process file: $_" "ERROR"
    }
}

# Modify the Write-EEEUOccurrencesToCSV function to include owner information
function Write-EEEUOccurrencesToCSV {
    param (
        [string]$filePath,
        [switch]$Append = $false,
        [array]$OccurrencesData = $global:EEEUOccurrences
    )
    try {
        # Create the file with headers if it doesn't exist or if we're not appending
        if (-not (Test-Path $filePath) -or -not $Append) {
            # Create empty file with headers - adding new columns for owner information
            "Url,ItemURL,RoleNames,OwnerName,OwnerEmail,CreatedDate" | Out-File -FilePath $filePath
        }

        # Group by URL, Item URL and Roles to remove duplicates
        $uniqueOccurrences = $OccurrencesData | 
        Group-Object -Property Url, ItemURL, RoleNames | 
        ForEach-Object { $_.Group[0] }
        
        # Append data to CSV
        foreach ($occurrence in $uniqueOccurrences) {
            # Manual CSV creation to handle special characters correctly
            $csvLine = "`"$($occurrence.Url)`",`"$($occurrence.ItemURL)`",`"$($occurrence.RoleNames)`",`"$($occurrence.OwnerName)`",`"$($occurrence.OwnerEmail)`",`"$($occurrence.CreatedDate)`""
            Add-Content -Path $filePath -Value $csvLine
        }
        
        Write-Log "EEEU occurrences have been written to $filePath"
    }
    catch {
        Write-Log "Failed to write EEEU occurrences to CSV file: $_" "ERROR"
    }
}

# Main script execution
$global:EEEUOccurrences = @()
$siteURLs = Read-SiteURLs -filePath $inputFilePath

# Create an empty output file with headers
Write-EEEUOccurrencesToCSV -filePath $outputFilePath

foreach ($siteURL in $siteURLs) {
    Write-Host "Looping through all files in $siteURL to locate EEEU in all files" -ForegroundColor Green
    # Clear the global collection for this site
    $global:EEEUOccurrences = @()
    
    if (Connect-SharePoint -siteURL $siteURL) {
        # Check connection success
        # Get all lists and libraries with throttling protection
        $lists = Invoke-WithRetry -ScriptBlock {
            Get-PnPList | Select-Object Title, id, Url | Where-Object { $_.Title -notin $ignoreFolders } | Select-Object -ExpandProperty Title
        }
        foreach ($list in $lists) {
            $allItems = Get-AllItemsInFolderAbs -siteURL $siteURL -folderUrl $list
            foreach ($item in $allItems) {
                Find-EEEUinFiles -item $item
            }
        }
        
        # Write the results for this site collection to the CSV
        if ($global:EEEUOccurrences.Count -gt 0) {
            Write-Host "Writing $($global:EEEUOccurrences.Count) EEEU occurrences from $siteURL to CSV..." -ForegroundColor Cyan
            Write-EEEUOccurrencesToCSV -filePath $outputFilePath -Append -OccurrencesData $global:EEEUOccurrences
        }
        else {
            Write-Host "No EEEU occurrences found in $siteURL" -ForegroundColor Green
            Write-Log "No EEEU occurrences found in $siteURL"
        }
    }
    
    Write-Host "Completed processing for $siteURL" -ForegroundColor Green
    Write-Log "Completed processing for $siteURL"
}

# Final message, don't need another CSV write since we've been writing after each site
Write-Host "EEEU occurrences scan completed. Results available in $outputFilePath" -ForegroundColor Green
Write-Log "EEEU occurrences scan completed. Results available in $outputFilePath"
