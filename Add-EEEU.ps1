<#
.SYNOPSIS
    Adds the "Everyone Except External Users" (EEEU) group with specified permissions to a file in SharePoint Online.

.DESCRIPTION
    This script connects to a SharePoint Online site using an Azure AD application authentication method (Client ID and Certificate Thumbprint).
    It retrieves a specified file from a SharePoint Online document library, breaks the permission inheritance on the file, and assigns the "Everyone Except External Users" group ("EEEU") with the specified permission level (default is "Read").

.PARAMETER appID
    The Azure AD Application (Client) ID used for authentication.

.PARAMETER thumbprint
    The certificate thumbprint associated with the Azure AD application for authentication.

.PARAMETER tenant
    The Azure AD Tenant ID.

.PARAMETER SiteURL
    The URL of the SharePoint Online site (OneDrive site) containing the target file.

.PARAMETER FilePath
    The relative path to the target file within the SharePoint Online document library.

.PARAMETER LoginName
    The login name representing the "Everyone Except External Users" group.

.PARAMETER logFilePath
    The path to the log file where script execution details are recorded.

.FUNCTIONS
    Write-Log
        Logs messages with timestamps and severity levels to the specified log file.

    Add-EEEUtoFile
        Breaks permission inheritance on the specified file and assigns the "Everyone Except External Users" group with the defined permission level.

.NOTES
    Ensure the Azure AD application has appropriate permissions to access and modify SharePoint Online resources.
    Requires the PnP.PowerShell module installed and configured.

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
    ./Add-EEEU.ps1
    Executes the script to add the "Everyone Except External Users" group with "Read" permissions to the specified file.

#>
#Tenant Level Information
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"

#path and file names
$SiteURL = "https://m365cpi13246019-my.sharepoint.com/personal/admin_m365cpi13246019_onmicrosoft_com"
$FilePath = "documents/ssssss.docx" # Path relative to the root of the OneDrive library
#$FilePath = "documents/files/book 1.xlsx" # Path relative to the root of the OneDrive library

#Script Parameters
$LoginName = "c:0-.f|rolemanager|spo-grid-all-users/$tenant"
$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "$env:TEMP\Add_EEEU_From_Files_$startime.txt"

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
try {
    Connect-PnPOnline -Url $SiteURL -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
    Write-Log "Connected to SharePoint Online"
}
catch {
    Write-Log "Failed to connect to SharePoint Online: $_" "ERROR"
    throw $_
}

# Get the file
try {
    $file = Get-PnPFile -Url $FilePath -AsListItem
    # Expand the ParentList property if needed
    if ($file) {
        Get-PnPProperty -ClientObject $file -Property ParentList
        # Write-Log "Expanded ParentList property for file: $($file.FieldValues.FileLeafRef)"
    }
    Write-Host "Retrieved file: $($file.FieldValues.FileLeafRef) on $SiteURL" -ForegroundColor Green
    Write-Log "Retrieved file: $($file.FieldValues.FileLeafRef) on $SiteURL"
}
catch {
    Write-Log "Failed to retrieve file: $_" "ERROR"
    throw $_
}

function Add-EEEUtoFile {
    param (
        [Microsoft.SharePoint.Client.ListItem]$file
    )
    try {
        # Break inheritance on the file
        $file.BreakRoleInheritance($true, $false)
        Write-Host "Broken Role Inheritance on File: $($file.FieldValues.FileLeafRef)" -ForegroundColor Yellow
        Write-Log "Broken Role Inheritance on File: $($file.FieldValues.FileLeafRef)"

        # Add EEEU to file
        #set role "Full Control", Design, Edit, Read, Contribute,Restricted view
        $role = "Read"
        Set-PnPListItemPermission -List $file.ParentList.Title -Identity $file.Id -AddRole $role -User $LoginName
        Write-Host "Added EEEU to File: $($file.FieldValues.FileLeafRef) with role $role" -ForegroundColor Cyan
        Write-Log "Added EEEU to File: $($file.FieldValues.FileLeafRef) with role $role"
    }
    catch {
        Write-Log "Failed to add EEEU to file: $_" "ERROR"
        throw $_
    }
}

# Add EEEU from file
Add-EEEUtoFile -file $file
write-host "Operations completed successfully" -ForegroundColor Green
Write-Log "Operations completed successfully"
