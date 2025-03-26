<#
.SYNOPSIS
    Script to remove Everyone Except External Users (EEEU) permissions from files listed in a CSV.

.DESCRIPTION
    This script processes a list of SharePoint files from a CSV file and removes the "Everyone except external users"
    permissions from each file. It uses PnP PowerShell with certificate-based authentication to connect to SharePoint
    Online sites and modify file permissions.

.PARAMETER None
    This script uses hardcoded values for authentication and file paths.

.NOTES
    File Name      : Remove-EEEUFromFileList.ps1
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
    
.INPUTS
    CSV file with columns for URL (site URL) and ItemURL (file path within the site)
    Located at: C:\temp\EEEUFileList.csv

.OUTPUTS
    Log file created in the user's temp directory with timestamp in the filename
    Console output with color-coded status messages

.EXAMPLE
    .\Remove-EEEUFromFileList.ps1

.FUNCTIONALITY
    1. Connects to SharePoint Online sites using app-based authentication
    2. Reads a CSV file containing site URLs and file paths
    3. For each file, retrieves its permissions
    4. Removes any permissions assigned to the "Everyone except external users" group
    5. Logs all operations to a log file
#>
# Tenant Level Information
$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"

# Script Parameters
$LoginName = "c:0-.f|rolemanager|spo-grid-all-users/$tenant"
$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "$env:TEMP\Remove_EEEU_From_File_List_$startime.txt"

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
function Connect-ToSharePoint {
    param (
        [string]$SiteURL
    )
    try {
        Connect-PnPOnline -Url $SiteURL -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
        Write-Log "Connected to SharePoint Online: $SiteURL"
    }
    catch {
        Write-Log "Failed to connect to SharePoint Online: $_" "ERROR"
        throw $_
    }
}

# Get the file
function Get-File {
    param (
        [string]$FilePath
    )
    try {
        $file = Get-PnPFile -Url $FilePath -AsListItem
        Write-Log "Retrieved file:  $($file.FieldValues.FileLeafRef) on $SiteURL"
        return $file
    }
    catch {
        Write-Log "Failed to retrieve file: $_" "ERROR"
        throw $_
    }
}

# Remove EEEU from file
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
                    Write-Host "Retrieved file: $($file.FieldValues.FileLeafRef) on $SiteURL" -ForegroundColor Yellow
                    Write-Log "Retrieved file:  $($file.FieldValues.FileLeafRef) on $SiteURL"
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

# Read the CSV file and process each row
$csvFilePath = "C:\temp\EEEUFileList.csv"
$csvData = Import-Csv -Path $csvFilePath
foreach ($row in $csvData) {

    Write-Host "Working on $FilePath in $siteURL to remove EEEU from all files" -ForegroundColor Green

    #The headings can be different in the CSV file, so we need to use the correct column names
    $SiteURL = $row.URL
    $FilePath = $row.ItemURL

    Connect-ToSharePoint -SiteURL $SiteURL
    
    $file = Get-File -FilePath $FilePath
    
    Remove-EEEUfromFile -file $file
}

Write-Log "Operations completed successfully"
Write-Host "Operations completed successfully"
