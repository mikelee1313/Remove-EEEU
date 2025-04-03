<#
.SYNOPSIS
    Removes specific permissions (EEEU) from a SharePoint Online file.

.DESCRIPTION
    This script connects to a SharePoint Online site using an Azure AD application identity and retrieves a specified file from a OneDrive library. It then removes permissions assigned to a specific user or group (identified by the login name) from the retrieved file. All actions and errors are logged to a timestamped log file located in the user's temporary directory.

.PARAMETER appID
    The Azure AD Application (Client) ID used for authentication.

.PARAMETER thumbprint
    The certificate thumbprint associated with the Azure AD application for authentication.

.PARAMETER tenant
    The Azure AD Tenant ID.

.PARAMETER SiteURL
    The URL of the SharePoint Online site (OneDrive) containing the target file.

.PARAMETER FilePath
    The relative path to the target file within the SharePoint Online document library.

.PARAMETER LoginName
    The login name of the user or group whose permissions will be removed from the file.

.OUTPUTS
    Logs detailed information and errors to a log file located at "$env:TEMP\Add_EEEU_From_File_<timestamp>.txt".

.NOTES
    Requires the PnP.PowerShell module.
    Ensure the Azure AD application has appropriate permissions to access and modify SharePoint Online resources.

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
    ./Remove-EEEU_from_File.ps1
    Executes the script with predefined parameters to remove permissions from the specified file.

#>
#Tenant Level Information
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"

#path and file names
$SiteURL = "https://m365cpi13246019-my.sharepoint.com/personal/admin_m365cpi13246019_onmicrosoft_com"
$FilePath = "/Documents/testdoc1.docx" # Path relative to the root of the OneDrive library

#Script Parameters
$LoginName = "c:0-.f|rolemanager|spo-grid-all-users/$tenant"
$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "$env:TEMP\Remove_EEEU_From_File_$startime.txt"


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
    Write-Host "Retrieved file: $($file.FieldValues.FileLeafRef) on $SiteURL" -ForegroundColor Green
    Write-Log "Retrieved file:  $($file.FieldValues.FileLeafRef) on $SiteURL"
}
catch {
    Write-Log "Failed to retrieve file: $_" "ERROR"
    throw $_
}

function Remove-EEEUfromFile {
    param (
        [Microsoft.SharePoint.Client.ListItem]$file
    )
    try {
        $RoleAssignment = @()
        $Permissions = Get-PnPProperty -ClientObject $file -Property RoleAssignments
        foreach ($RoleAssignment in $Permissions) {
            Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
            if ($RoleAssignment.Member.LoginName -eq $LoginName) {
                $roleuser = $RoleAssignment.Member.LoginName
                $rolelevel = $RoleAssignment.RoleDefinitionBindings
                foreach ($role in $rolelevel) {
                    Set-PnPListItemPermission -List "Documents" -Identity $file.Id -RemoveRole $role.Name -User $roleuser
                    Write-Host "Removed Role: $($role.Name) from User: $($roleuser) in File: $($file.FieldValues.FileLeafRef)" -ForegroundColor Yellow
                    Write-Log "Removed Role: $($role.Name) from User: $($roleuser) in File: $($file.FieldValues.FileLeafRef)"
                }
            }
        }
    }
    catch {
        Write-Log "Failed to remove EEEU from file: $_" "ERROR"
        throw $_
    }
}


# Remove EEEU from  file
Remove-EEEUfromFile -file $file
Write-Host "Operations completed successfully"
Write-Log "Operations completed successfully"
