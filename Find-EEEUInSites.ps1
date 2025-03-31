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
    Date Created   : 3/31/2025

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
$LoginName = "c:0-.f|rolemanager|spo-grid-all-users/$tenant"
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
        Connect-PnPOnline -Url $siteURL -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
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
            $items = Get-PnPListItem -List $folderUrl -PageSize 500
            $allItems += $items | Where-Object { $_["FileLeafRef"] -like "*.*" }

            $subFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $folderUrl -ItemType Folder
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
        #Write-Log "Failed to retrieve items from folder $folderUrl : $_" "ERROR"
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
       
        $file = Get-PnPFile -Url $fileUrl -AsListItem

        $Permissions = Get-PnPProperty -ClientObject $file -Property RoleAssignments

        if ($Permissions) {
            $roles = @()
            foreach ($RoleAssignment in $Permissions) {
                Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
                #Write-Log "Checking  File: $($fileUrl) role assignment: $($RoleAssignment.Member.LoginName), Role: $($RoleAssignment.RoleDefinitionBindings.name)"

                if ($RoleAssignment.Member.LoginName -eq $LoginName -and $RoleAssignment.RoleDefinitionBindings.name -ne 'Limited Access') {
                    $rolelevel = $RoleAssignment.RoleDefinitionBindings
                    foreach ($role in $rolelevel) {
                        $roles += $role.Name
                    }   
                }
            }
            if ($roles.Count -gt 0) {
                $global:EEEUOccurrences += [PSCustomObject]@{
                    Url   = $SiteURL
                    ItemURL  = $file.FieldValues.FileRef
                    RoleNames = ($roles -join ", ")
                }
                Write-Host "Located EEEU in file: $($file.FieldValues.FileLeafRef) on $SiteURL" -ForegroundColor Green
                Write-Log "Located EEEU in file: $($file.FieldValues.FileLeafRef) on $SiteURL"
            }
        }
    }
    catch {
        Write-Log "Failed to process file: $_" "ERROR"
    }
}


# Write EEEU occurrences to CSV file (remove duplicates)
function Write-EEEUOccurrencesToCSV {
    param (
        [string]$filePath
    )
    try {
        $global:EEEUOccurrences | Group-Object SiteURL, FileName, RoleNames | ForEach-Object { $_.Group } | Export-Csv -Path $filePath -NoTypeInformation
        Write-Log "EEEU occurrences have been written to $filePath"
    }
    catch {
        Write-Log "Failed to write EEEU occurrences to CSV file: $_" "ERROR"
    }
}

# Main script execution
$global:EEEUOccurrences = @()
$siteURLs = Read-SiteURLs -filePath $inputFilePath

foreach ($siteURL in $siteURLs) {
    Write-Host "Looping through all files in $siteURL to locate EEEU in all files" -ForegroundColor Green

    if (Connect-SharePoint -siteURL $siteURL) {
        # Check connection success
        #Get all lists and libraries
        #$lists = Get-PnPList | Where-Object { $_.BaseTemplate -eq "DocumentLibrary" -or $_.BaseTemplate -eq "GenericList" } | Select-Object -ExpandProperty RootFolder.ServerRelativeUrl
        $lists = Get-PnPList | Select-Object Title, id, Url | Where-Object { $_.Title -notin $ignoreFolders } | Select-Object -ExpandProperty Title
        foreach ($list in $lists) {
            $allItems = Get-AllItemsInFolderAbs -siteURL $siteURL -folderUrl $list
            foreach ($item in $allItems) {
                Find-EEEUinFiles -item $item
            }
        }
    }
    Write-Host "Operations completed successfully"
    Write-Log "Operations completed successfully"
}

Write-EEEUOccurrencesToCSV -filePath $outputFilePath
Write-Host "EEEU occurrences have been located and reported in $outputFilePath" -ForegroundColor Green
Write-Log "EEEU occurrences have been located and reported in $outputFilePath"
