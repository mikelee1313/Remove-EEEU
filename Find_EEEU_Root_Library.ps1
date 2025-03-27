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
#>

$appID = "1e488dc4-1977-48ef-8d4d-9856f4e04536"
$thumbprint = "5EAD7303A5C7E27DB4245878AD554642940BA082"
$tenant = "9cfc42cb-51da-4055-87e9-b20a170b6ba3"
$spadmin = "m365cpi13246019"

$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "$env:TEMP\Find_EEEU_Root_Library_$startime.txt"

function Write-Log {
    param (
        [string]$message,
        [string]$level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $level - $message"
    Add-Content -Path $logFilePath -Value $logMessage
}

# requires SharePoint > Application > Sites.FullControl.All
Connect-PnPOnline -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant -Url "https://$spadmin-admin.sharepoint.com" -ErrorAction Stop

# pull all sites and filter out non-onedrive sites
$personalSites = Get-PnPTenantSite -IncludeOneDriveSites | Where-Object -Property Template -match "SPSPERS"

# Function to check for EEEU at the default document library level
function Test-EEEUDocumentLibrary {
    param (
        [string]$siteUrl
    )

    Write-Host "[$(Get-Date)] - Checking default document library for EEEU: $siteUrl"
    Write-Log "[$(Get-Date)] - Checking default document library for EEEU: $siteUrl"

    try {
        # Connect to the site
        Connect-PnPOnline -Url $siteUrl -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant -ErrorAction Stop

        # Get the default document library
        $documentLibrary = Get-PnPList -Identity "documents"

        # Get RoleAssignments for the document library
        $roleAssignments = Get-PnPProperty -ClientObject $documentLibrary -Property RoleAssignments

        # Enumerate RoleAssignments
        foreach ( $roleAssignment in $roleAssignments ) {
            # Pull principal LoginName
            $null = Get-PnPProperty -ClientObject $roleAssignment.Member -Property LoginName

            # Enumerate principals
            foreach ( $member in $roleAssignment.Member ) {
                # Check if principal is EEEU
                if ( $member.LoginName -match "spo-grid-all-users" ) {
                    # Hydrate RoleDefinitionBindings (permissions)
                    $null = Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings

                    # Filter out hidden permissions
                    $roleDefinitionBindings = $roleAssignment.RoleDefinitionBindings | Where-Object -Property Hidden -eq $false
                    if ( $roleDefinitionBindings ) {
                        Write-Host "Found EEEU in Document Library: $($member.LoginName) on $siteUrl" -ForegroundColor Red
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
        Write-Host "Error processing document library: $($_)" -ForegroundColor Red
    }
}

# enumerate onedrive sites
$results = foreach ( $personalSite in $personalSites ) {
    Write-Host "Checking Root for EEEU on: $($personalSite.Url)"
    Write-Log "Checking Root for EEEU on: $($personalSite.Url)"

    try {
        # connect to the onedrive site
        Connect-PnPOnline -Url $personalSite.Url -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant -ErrorAction Stop

        # get the rootweb + RoleAssignments
        $web = Get-PnPWeb -Includes RoleAssignments

        # enumerate RoleAssignments
        foreach ( $roleAssignment in $web.RoleAssignments ) {
            # pull principal LoginName
            $null = Get-PnPProperty -ClientObject $roleAssignment.Member -Property LoginName

            # enumerate principals
            foreach ( $member in $roleAssignment.Member ) {
                # check if principal is EEEU
                if ( $member.LoginName -match "spo-grid-all-users" ) {
                    # hydrate RoleDefinitionBindings (permissions)
                    $null = Get-PnPProperty -ClientObject $roleAssignment -Property RoleDefinitionBindings

                    # filter out hidden permissions
                    $roleDefinitionBindings = $roleAssignment.RoleDefinitionBindings | Where-Object -Property Hidden -eq $false
                    if ( $roleDefinitionBindings ) {
                        write-host "Found EEEU: $($member.LoginName) on $($personalSite.Url)" -ForegroundColor Red
                        Write-Log "Found EEEU: $($member.LoginName) on $($personalSite.Url)" -level "WARNING"
                        
                        # output object
                        [PSCustomObject] @{
                            SiteUrl                = $personalSite.Url
                            Owner                  = $personalSite.Owner
                            Claim                  = $member.LoginName
                            RoleDefinitionBindings = $roleDefinitionBindings.Name -join ","
                            LibraryName            = "RootWeb"
                        }
                    }
                    else { 
                        continue 
                    }
                }
            
            }
        }

        # Call the new function to check for EEEU at the default document library level
        Test-EEEUDocumentLibrary -siteUrl $personalSite.Url

    }
    catch {
        Write-Host "Error processing site: $($_)" -ForegroundColor Red
    }
}

$results | Export-Csv -Path "$env:TEMP\OneDrive_EEEU_Root_Library_$startime.csv" -NoTypeInformation
