<#
.SYNOPSIS
    Script to add Everyone Except External Users (EEEU) permissions to a SharePoint site and its document library.

    Note: This is only for testing purposes and should only be used to test the side effects of adding EEEU permissions to a OneDrive site and document library.

.DESCRIPTION
    This script connects to a SharePoint Online site using certificate-based authentication and assigns Read permissions 
    to the "Everyone Except External Users" group at both the site level and document library level.
    
    The script includes logging functionality to track operations and errors to a timestamped log file in the temp directory.

.PARAMETER appID
    The Azure AD application ID used for certificate-based authentication.

.PARAMETER thumbprint
    The certificate thumbprint used for authentication.

.PARAMETER tenant
    The tenant ID for the Microsoft 365 environment.

.PARAMETER SiteURL
    The URL of the SharePoint site to modify permissions.

.PARAMETER LoginName
    The claim identity for the "Everyone Except External Users" group.

.NOTES
    File Name      : Add-EEEUtoOneDrive-RootandLibrary.ps1
    Author         : Mike Lee
    Date           : 3/27/25
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
    .\Add-EEEUPermissions.ps1
#>

$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"

# Path and file names
$SiteURL = "https://m365cpi13246019-my.sharepoint.com/personal/admin_m365cpi13246019_onmicrosoft_com"

# Script Parameters
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

function Add-EEEUtoFile {
    try {
        # Add EEEU to Site Level for testing
        $role = "Read"
        Set-PnPWebPermission -AddRole $role -User $LoginName
        Write-Host "Added EEEU to site level: $($SiteURL) with role $role" -ForegroundColor Cyan
        Write-Log "Added EEEU to site level: $($SiteURL) with role $role"
        
        # Add EEEU to default document library level
        $documentLibrary = Get-PnPList -Identity "Documents"
        Set-PnPListPermission -Identity $documentLibrary.Title -AddRole $role -User $LoginName 
        Write-Host "Added EEEU to default document library level: $documentLibrary with role $role" -ForegroundColor Cyan
        Write-Log "Added EEEU to default document library level: $documentLibrary with role $role"
    }
    catch {
        Write-Log "Failed to add EEEU to file: $_" "ERROR"
        throw $_
    }
}

# Add EEEU from file
Add-EEEUtoFile
Write-Host "Operations completed successfully" -ForegroundColor Green
Write-Log "Operations completed successfully"
