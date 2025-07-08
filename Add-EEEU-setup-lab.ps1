#Tenant Level Information
$appID = "5baa1427-1e90-4501-831d-a8e67465f0d9"                 # This is your Entra App ID
$thumbprint = "B696FDCFE1453F3FBC6031F54DE988DA0ED905A9"        # This is certificate thumbprint
$tenant = "85612ccb-4c28-4a34-88df-a538cc139a51"                # This is your Tenant ID

#Script Parameters
$LoginName = "c:0-.f|rolemanager|spo-grid-all-users/$tenant"
$startime = Get-Date -Format "yyyyMMdd_HHmmss"
$logFilePath = "$env:TEMP\Labsetup_Add_EEEU_To_Files_$startime.txt"

# Define an array of site URLs and file paths
$filesToProcess = @(
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/GlobalMarketing"; FilePath = "/Shared%20Documents/GlobalMarketingDoc1.docx" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/GlobalSales"; FilePath = "/Lists/TestList1/1_.000" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/Leadership"; FilePath = "/Shared%20Documents/Executive FY Goals.pptx" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/leadership-connection"; FilePath = "/Shared%20Documents/Folder1/ImportantNumbers.xlsx" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/Mark8ProjectTeam"; FilePath = "/Shared%20Documents/Mark8-MarketingCampaign.docx" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/collabnogroup"; FilePath = "/Shared%20Documents/doc123.docx" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/AdminsTeam"; FilePath = "/Shared%20Documents/ch1/share.docx" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/commsite1" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/commsite1"; FilePath = "/Shared%20Documents" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/commsite1"; FilePath = "/Lists/List1" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/commsite1"; FilePath = "/doc1/thisisafolder" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/commsite1/sub1"; FilePath = "/Shared%20Documents/ProjectX.pptx" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/commsite1"; FilePath = "/doc1/groups.csv" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/commsite1"; FilePath = "/Shared%20Documents/SiteSharingSettings_20250522_184414.csv" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/commsite1"; FilePath = "/Shared%20Documents/SiteSharingSettings_20250522_184513.csv" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/commsite1"; FilePath = "/Shared%20Documents/SiteSharingSettings_20250522_184618.csv" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/commsite1"; FilePath = "/Shared%20Documents/SiteSharingSettings_20250522_184850.csv" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/commsite1"; FilePath = "/Shared%20Documents/SiteSharingSettings_20250522_185108.csv" },
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/commsite1"; FilePath = "/Shared%20Documents/SiteSharingSettings_20250522_190005.csv" },   
    @{ SiteURL = "https://m365x61250205.sharepoint.com/sites/TestSiteGroupConnected/sub2/sub3"; FilePath = "/doclib2/OversharedDoc1.docx" },
    @{ SiteURL = "https://m365x61250205-my.sharepoint.com/personal/admin_m365x61250205_onmicrosoft_com"; FilePath = "/documents/Finance.pbix" },
    @{ SiteURL = "https://m365x61250205-my.sharepoint.com/personal/admin_m365x61250205_onmicrosoft_com"; FilePath = "/documents/attachments/Proposed_agenda_topics.docx" },
    @{ SiteURL = "https://m365x61250205-my.sharepoint.com/personal/alland_m365x61250205_onmicrosoft_com"; FilePath = "/documents/Personal Info/To Do list Friday.xlsx" }
)

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

# Function to add EEEU permissions to a file
function Add-EEEUtoFile {
    param (
        [Microsoft.SharePoint.Client.ListItem]$file,
        [string]$listName = "Documents"
    )
    try {
        # Break inheritance on the file
        $file.BreakRoleInheritance($true, $false)
        Write-Host "Broken Role Inheritance on File: $($file.FieldValues.FileLeafRef)" -ForegroundColor Yellow
        Write-Log "Broken Role Inheritance on File: $($file.FieldValues.FileLeafRef)"

        # Add EEEU to file
        #set role "Full Control", Design, Edit, Read, Contribute,Restricted view
        $role = "Edit"
        Set-PnPListItemPermission -List $file.ParentList.Title -Identity $file.Id -AddRole $role -User $LoginName
        Write-Host "Added EEEU to File: $($file.FieldValues.FileLeafRef) with role $role" -ForegroundColor Cyan
        Write-Log "Added EEEU to File: $($file.FieldValues.FileLeafRef) with role $role"
    }
    catch {
        Write-Log "Failed to add EEEU to file: $_" "ERROR"
        throw $_
    }
}

# Function to add EEEU permissions to a document library
function Add-EEEUtoLibrary {
    param (
        [string]$libraryName
    )
    try {
        # Get the library
        $library = Get-PnPList -Identity $libraryName -ErrorAction Stop
        
        # Break inheritance on the library
        $library.BreakRoleInheritance($true, $false)
        $library.Update()
        Invoke-PnPQuery
        Write-Host "Broken Role Inheritance on Library: $libraryName" -ForegroundColor Yellow
        Write-Log "Broken Role Inheritance on Library: $libraryName"

        # Add EEEU to library
        $role = "Edit"
        Set-PnPListPermission -Identity $libraryName -AddRole $role -User $LoginName
        Write-Host "Added EEEU to Library: $libraryName with role $role" -ForegroundColor Cyan
        Write-Log "Added EEEU to Library: $libraryName with role $role"
    }
    catch {
        Write-Log "Failed to add EEEU to library: $_" "ERROR"
        throw $_
    }
}

# Function to add EEEU permissions at the web (site) level
function Add-EEEUtoWeb {
    try {
        # Get the current web
        $web = Get-PnPWeb -ErrorAction Stop
        
        # Break inheritance on the web if it's not already broken
        if ($web.HasUniqueRoleAssignments -eq $false) {
            $web.BreakRoleInheritance($true, $true)
            $web.Update()
            Invoke-PnPQuery
            Write-Host "Broken Role Inheritance on Web: $($web.Title)" -ForegroundColor Yellow
            Write-Log "Broken Role Inheritance on Web: $($web.Title)"
        }
        else {
            Write-Host "Web already has unique permissions: $($web.Title)" -ForegroundColor Yellow
            Write-Log "Web already has unique permissions: $($web.Title)"
        }

        # Add EEEU to web
        $role = "Edit"
        Set-PnPWebPermission -User $LoginName -AddRole $role
        Write-Host "Added EEEU to Web: $($web.Title) with role $role" -ForegroundColor Cyan
        Write-Log "Added EEEU to Web: $($web.Title) with role $role"
    }
    catch {
        Write-Log "Failed to add EEEU to web: $_" "ERROR"
        throw $_
    }
}

# Function to add EEEU permissions to a folder
function Add-EEEUtoFolder {
    param (
        [string]$folderPath,
        [string]$libraryName
    )
    try {
        Write-Host "Getting folder: $folderPath in library: $libraryName" -ForegroundColor Yellow
        Write-Log "Getting folder: $folderPath in library: $libraryName"
        
        # First get the list/library
        $list = Get-PnPList -Identity $libraryName -ErrorAction Stop
        
        # Extract the folder's relative URL from the full path
        $folderRelativeUrl = $folderPath
        # If folder path starts with the library name, remove it to get relative path within the library
        if ($folderRelativeUrl.StartsWith("/$libraryName/")) {
            $folderRelativeUrl = $folderRelativeUrl.Substring(("/$libraryName/").Length)
        }
        
        # Get the folder using server relative URL
        $folder = Get-PnPFolder -Url $folderPath -ErrorAction Stop
        
        # Get the folder as a list item (more reliable for permission setting)
        $folderItem = Get-PnPListItem -List $list -Query "<View><Query><Where><Eq><FieldRef Name='FileRef'/><Value Type='Text'>$folderPath</Value></Eq></Where></Query></View>" -ErrorAction Stop
        
        if ($null -eq $folderItem) {
            # Try alternate approach to get the folder item
            $serverRelativeUrl = $folder.ServerRelativeUrl
            $folderItem = Get-PnPListItem -List $list -Query "<View><Query><Where><Eq><FieldRef Name='FileRef'/><Value Type='Text'>$serverRelativeUrl</Value></Eq></Where></Query></View>" -ErrorAction Stop
        }
        
        if ($null -eq $folderItem) {
            # If still null, try another approach by getting all items and filtering
            Write-Host "Using alternative method to find folder item..." -ForegroundColor Yellow
            Write-Log "Using alternative method to find folder item..."
            
            # Get all folders in the library
            $allItems = Get-PnPListItem -List $list -PageSize 2000 | Where-Object { 
                $_.FieldValues.FileRef -eq $folderPath -or $_.FieldValues.FileRef -eq $folder.ServerRelativeUrl 
            }
            
            if ($allItems -and $allItems.Count -gt 0) {
                $folderItem = $allItems[0]
            }
        }
        
        if ($null -eq $folderItem) {
            throw "Could not find folder as list item: $folderPath"
        }
        
        Write-Host "Found folder item ID: $($folderItem.Id) for path: $folderPath" -ForegroundColor Green
        Write-Log "Found folder item ID: $($folderItem.Id) for path: $folderPath"
        
        # Break inheritance on the folder
        $folderItem.BreakRoleInheritance($true, $false)
        $folderItem.Context.ExecuteQuery()
        Write-Host "Broken Role Inheritance on Folder: $folderPath" -ForegroundColor Yellow
        Write-Log "Broken Role Inheritance on Folder: $folderPath"

        # Add EEEU to folder
        $role = "Edit"
        Set-PnPListItemPermission -List $libraryName -Identity $folderItem.Id -AddRole $role -User $LoginName
        Write-Host "Added EEEU to Folder: $folderPath with role $role" -ForegroundColor Cyan
        Write-Log "Added EEEU to Folder: $folderPath with role $role"
    }
    catch {
        Write-Log "Failed to add EEEU to folder: $_" "ERROR"
        throw $_
    }
}

# Process each file
Write-Log "Starting file processing"

# Variable to track current site URL for efficient connections
$currentSiteURL = ""

foreach ($item in $filesToProcess) {
    $SiteURL = $item.SiteURL
    $FilePath = $item.FilePath

    Write-Log "Processing item on site: $SiteURL, Path: $FilePath"

    try {
        # Only connect if we're working with a different site
        if ($currentSiteURL -ne $SiteURL) {
            Connect-PnPOnline -Url $SiteURL -ClientId $appID -Thumbprint $thumbprint -Tenant $tenant
            Write-Log "Connected to SharePoint Online site: $SiteURL"
            $currentSiteURL = $SiteURL
        }
        
        # Check if this is a web-level operation (no FilePath)
        if ([string]::IsNullOrEmpty($FilePath)) {
            # Handle web-level permissions
            Write-Host "Processing web-level permissions for site: $SiteURL" -ForegroundColor Green
            Write-Log "Processing web-level permissions for site: $SiteURL"
            Add-EEEUtoWeb
            Write-Host "Successfully processed web-level permissions for: $SiteURL" -ForegroundColor Green
            Write-Log "Successfully processed web-level permissions for: $SiteURL"
            continue
        }
        
        # Determine if this is a library/list or a file
        $isLibrary = $false
        $isFolder = $false
        $libraryName = "Documents" # Default
        
        # Special handling for paths without leading slash (typical for Lists)
        if (-not $FilePath.StartsWith("/")) {
            $isLibrary = $true
            $libraryName = $FilePath
            Write-Host "Processing list without leading slash: $libraryName" -ForegroundColor Yellow
            Write-Log "Processing list without leading slash: $libraryName"
        }
        # Regular path handling (with leading slash)
        elseif ($FilePath -match "^/([^/]+)$" -or $FilePath -match "^/([^/]+)/$") {
            # This is likely a library path
            $isLibrary = $true
            $libraryName = $Matches[1] -replace "%20", " "
            if ($libraryName -eq "Shared Documents") {
                $libraryName = "Documents"
            }
        } 
        # Lists path with leading slash
        elseif ($FilePath -match "^/Lists/([^/]+)$") {
            $isLibrary = $true
            $libraryName = "Lists/" + $Matches[1]
            Write-Host "Processing list with path: $libraryName" -ForegroundColor Yellow
            Write-Log "Processing list with path: $libraryName"
        }
        # Check if this is a folder (no file extension and contains subdirectories)
        elseif ($FilePath -notmatch "\.[^.]+$" -or $FilePath -match "/[^/.]+$") {
            # Check if the path exists and is a folder
            try {
                Write-Host "Checking if path is a folder: $FilePath" -ForegroundColor Yellow
                Write-Log "Checking if path is a folder: $FilePath"
                
                # Determine the library name from the path
                $pathParts = $FilePath -split "/"
                if ($pathParts.Length -gt 1) {
                    $libraryName = $pathParts[1] -replace "%20", " "
                    if ($libraryName -eq "Shared Documents") {
                        $libraryName = "Documents"
                    }
                    
                    # Try to get the folder to check if it exists
                    $folderExists = $false
                    try {
                        $folder = Get-PnPFolder -Url $FilePath -ErrorAction SilentlyContinue
                        if ($folder -ne $null) {
                            $folderExists = $true
                            Write-Host "Confirmed path is a folder: $FilePath" -ForegroundColor Green
                            Write-Log "Confirmed path is a folder: $FilePath"
                        }
                    }
                    catch {
                        Write-Host "Path is not a folder: $FilePath - $_" -ForegroundColor Yellow
                        Write-Log "Path is not a folder: $FilePath - $_" "WARNING"
                    }
                    
                    if ($folderExists) {
                        $isFolder = $true
                        Write-Host "Detected folder path: $FilePath in library: $libraryName" -ForegroundColor Yellow
                        Write-Log "Detected folder path: $FilePath in library: $libraryName"
                    }
                }
            }
            catch {
                # Not a folder, continue with file processing
                $isFolder = $false
                Write-Host "Error checking if path is folder: $FilePath - $_" -ForegroundColor Yellow
                Write-Log "Error checking if path is folder: $FilePath - $_" "WARNING"
            }
        }
        else {
            # This is a file, determine list name from file path
            if ($FilePath -match "/([^/]+)/") {
                $pathParts = $FilePath -split "/"
                if ($pathParts.Length -gt 1) {
                    $libraryName = $pathParts[1] -replace "%20", " "
                    # Handle special cases
                    if ($libraryName -eq "Shared Documents") {
                        $libraryName = "Documents"
                    }
                }
            }
        }
        
        if ($isLibrary) {
            # Handle library/list permissions
            Write-Host "Processing library/list: $libraryName on $SiteURL" -ForegroundColor Green
            Write-Log "Processing library/list: $libraryName on $SiteURL"
            Add-EEEUtoLibrary -libraryName $libraryName
            Write-Host "Successfully processed library/list: $libraryName" -ForegroundColor Green
            Write-Log "Successfully processed library/list: $libraryName"
        }
        elseif ($isFolder) {
            # Handle folder permissions
            Write-Host "Processing folder: $FilePath on $SiteURL" -ForegroundColor Green
            Write-Log "Processing folder: $FilePath on $SiteURL"
            Add-EEEUtoFolder -folderPath $FilePath -libraryName $libraryName
            Write-Host "Successfully processed folder: $FilePath" -ForegroundColor Green
            Write-Log "Successfully processed folder: $FilePath"
        }
        else {
            # Handle file permissions
            $file = Get-PnPFile -Url $FilePath -AsListItem
            if ($file) {
                Get-PnPProperty -ClientObject $file -Property ParentList
            }
            Write-Host "Retrieved file: $($file.FieldValues.FileLeafRef) on $SiteURL" -ForegroundColor Green
            Write-Log "Retrieved file: $($file.FieldValues.FileLeafRef) on $SiteURL"

            # Add EEEU to file
            Add-EEEUtoFile -file $file -listName $libraryName
            Write-Host "Successfully processed file: $($file.FieldValues.FileLeafRef)" -ForegroundColor Green
            Write-Log "Successfully processed file: $($file.FieldValues.FileLeafRef)"
        }
    }
    catch {
        Write-Host "Error processing item on site: $SiteURL, Path: $FilePath - $_" -ForegroundColor Red
        Write-Log "Error processing item on site: $SiteURL, Path: $FilePath - $_" "ERROR"
    }
}

Write-Host "All operations completed successfully" -ForegroundColor Green
Write-Log "All operations completed successfully"
