<#
.SYNOPSIS
    Find and Removes specific user permissions ("EEEU") from files in a Onerive Site.

.DESCRIPTION
    This script connects to a specified SharePoint Online site using application credentials (Client ID and certificate thumbprint).
    It recursively traverses through the specified document library, excluding predefined folders, and removes permissions associated
    with a specific user or group (defined by $LoginName) from each file. Permissions are removed only if they are not "Limited Access".

.PARAMETER appID
    The Client ID of the Azure AD application used for authentication.

.PARAMETER thumbprint
    The thumbprint of the certificate associated with the Azure AD application.

.PARAMETER tenant
    The Azure AD tenant ID.

.PARAMETER LoginName
    The login name of the user or group whose permissions will be removed from files.

.PARAMETER SiteURL
    The URL of the SharePoint Online site containing the document library.

.PARAMETER DocumentLibrary
    The name of the document library to process.

.PARAMETER ignoreFolders
    An array of folder names to exclude from processing.

.PARAMETER BatchSize
    The number of items to retrieve per batch when querying SharePoint.

.OUTPUTS
    Generates a log file at the specified temporary directory ($env:TEMP) with details of operations performed.

.NOTES
    Ensure the PnP.PowerShell module is installed and properly configured.
    The script requires appropriate permissions to modify file permissions in SharePoint Online.

    Authors: Mike Lee
    Date: 3/25/2025

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
    ./Find_and_Remove_EEEU_From_Files_in_OneDrive.ps1

    Executes the script with predefined parameters, connects to SharePoint Online, and removes specified user permissions from files.

#>
# Tenant Level Information
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"

# Script Parameters
$LoginName = "c:0-.f|rolemanager|spo-grid-all-users/$tenant"
$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "$env:TEMP\Remove_EEEU_From_Files_$startime.txt"

# Path and file names
$SiteURL = "https://m365cpi13246019-my.sharepoint.com/personal/admin_m365cpi13246019_onmicrosoft_com"
$DocumentLibrary = "Documents"

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

# Connect to SharePoint Online
function Connect-SharePoint {
    try {
        Connect-PnPOnline -Url $SiteURL -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
        Write-Log "Connected to SharePoint Online"
    }
    catch {
        Write-Log "Failed to connect to SharePoint Online: $_" "ERROR"
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
                    Get-ListItems -BatchSize $BatchSize -FolderPath $folder.ServerRelativeUrl #Pass FolderPath
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
Write-Host "Looping thorugh all files in $SiteURL to remove EEEU from all files" -ForegroundColor Green

Connect-SharePoint
$BatchSize = 5 # Adjust the batch size as needed
Get-ListItems -BatchSize $BatchSize -FolderPath "" # Added FolderPath as "" for the root of the document library.

Write-Host "Operations completed successfully"
Write-Log "Operations completed successfully"
