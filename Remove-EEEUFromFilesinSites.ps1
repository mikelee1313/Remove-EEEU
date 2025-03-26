<#
.SYNOPSIS
    Removes "Everyone except external users" (EEEU) permissions from files using a list of SharePoint Online and OneDrive sites.

.DESCRIPTION
    This script scans through SharePoint Online / OneDrive sites provided in an input file and removes the "Everyone except external users"
    permissions from all files. It recursively processes all folders and subfolders in document libraries, 
    skipping a predefined list of system folders.

.PARAMETER appID
    The application ID for the Azure AD app registration.

.PARAMETER thumbprint
    The certificate thumbprint used for authentication.

.PARAMETER tenant
    The tenant ID of the Microsoft 365 tenant.

.PARAMETER LoginName
    The login name of the "Everyone except external users" group.

.PARAMETER inputFilePath
    The path to a text file containing SharePoint site URLs to process.

.NOTES
    File Name      : Remove-EEEUFromFilesinSites.ps1
    Author         : Mike Lee
    Date           : 3/26/25

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
    .\Remove-EEEUFromFilesinSites.ps1

.FUNCTIONALITY
    - Connects to SharePoint Online sites using certificate-based authentication
    - Recursively traverses document libraries and folders
    - Identifies and removes EEEU permissions from files
    - Logs all operations to a timestamped log file in the %TEMP% directory
    - Skips specified system folders to avoid errors and improve performance
#>
# Tenant Level Information
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"

# Script Parameters
$LoginName = "c:0-.f|rolemanager|spo-grid-all-users/$tenant"
$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "$env:TEMP\Remove_EEEU_From_Files_in_Sites_$startime.txt"

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
    }
    catch {
        Write-Log "Failed to connect to SharePoint Online at $siteURL : $($_.Exception.Message)" "ERROR"
        throw $_
    }
}

# List of folders to ignore
$ignoreFolders = @(
    "_catalogs",
    "_cts",
    "_private",
    "_vti_pvt",
    "images",
    "Lists",
    "Reference 778a30bb4f074ae3bec315889ee34b88",
    "Sharing Links",
    "SitePages",
    "Social"
)

# Get list items with batch size
function Get-ListItems {
    param (
        [int]$BatchSize,
        [string]$DocumentLibrary,
        [string]$FolderPath = "" # Ensure FolderPath is passed and defined.
    )
    try {
        if ($FolderPath -eq "") {
            $ListItems = Get-PnPListItem -List $DocumentLibrary -PageSize $BatchSize | Where-Object { $_["FileLeafRef"] -like "*.*" }
            Write-Log "Retrieved $($ListItems.Count) items from $DocumentLibrary (root)"
        }
        else {
            $ListItems = Get-PnPListItem -List $DocumentLibrary -PageSize $BatchSize -FolderServerRelativeUrl $FolderPath | Where-Object { $_["FileLeafRef"] -like "*.*" }
            Write-Log "Retrieved $($ListItems.Count) items from $DocumentLibrary at $FolderPath"
        }

        foreach ($item in $ListItems) {
            Remove-EEEUfromFiles -item $item
        }

        # Retrieve subfolders only if there are items in the current folder
        if ($ListItems.Count -gt 0) {
            if ($FolderPath -eq "") {
                $SubFolders = Get-PnPFolderItem -FolderSiteRelativeUrl "/" -ItemType Folder # use root folder relative url
            }
            else {
                $SubFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderPath -ItemType Folder
            }

            foreach ($folder in $SubFolders) {
                if ($folder.Name -notin $ignoreFolders) {
                    Get-ListItems -BatchSize $BatchSize -DocumentLibrary $DocumentLibrary -FolderPath $folder.ServerRelativeUrl #Pass FolderPath
                }
                else {
                    Write-Log "Ignoring folder: $($folder.Name)"
                }
            }
        }
    }
    catch {
        Write-Log "Failed to retrieve list items: $_" "ERROR"
        throw $_
    }
}

# Process file
function Remove-EEEUfromFiles {
    param (
        $item
    )
    try {
        $file = Get-PnPFile -Url $item.FieldValues.FileRef -AsListItem
        $Permissions = Get-PnPProperty -ClientObject $file -Property RoleAssignments
        if ($Permissions) {
            foreach ($RoleAssignment in $Permissions) {
                Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
                if ($RoleAssignment.Member.LoginName -eq $LoginName -and $RoleAssignment.RoleDefinitionBindings.name -ne 'Limited Access') {
                    $roleuser = $RoleAssignment.Member.LoginName
                    $rolelevel = $RoleAssignment.RoleDefinitionBindings
                    foreach ($role in $rolelevel) {
                        Write-Host "Retrieved file: $($file.FieldValues.FileLeafRef) on $SiteURL" -ForegroundColor Green
                        Write-Log "Retrieved file:  $($file.FieldValues.FileLeafRef) on $SiteURL"
                        Set-PnPListItemPermission -List $DocumentLibrary -Identity $file.Id -RemoveRole $role.Name -User $roleuser
                        Write-Host "Removed Role: $($role.Name) from User: $($roleuser) in File: $($file.FieldValues.FileLeafRef)" -ForegroundColor Yellow
                        Write-Log "Removed Role: $($role.Name) from User: $($roleuser) in File: $($file.FieldValues.FileLeafRef)"
                    }
                }
            }
        }
    }
    catch {
        Write-Log "Failed to process file: $_" "ERROR"
        throw $_
    }
}

# Main script execution
$siteURLs = Read-SiteURLs -filePath $inputFilePath

foreach ($siteURL in $siteURLs) {
    Write-Host "Looping through all files in $siteURL to remove EEEU from all files" -ForegroundColor Green

    Connect-SharePoint -siteURL $siteURL
    $BatchSize = 5 # Adjust the batch size as needed
    Get-ListItems -BatchSize $BatchSize -DocumentLibrary "Documents" -FolderPath "" # Added FolderPath as "" for the root of the document library.

    Write-Host "Operations completed successfully"
    Write-Log "Operations completed successfully"
}
