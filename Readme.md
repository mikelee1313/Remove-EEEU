Certainly! Here's a draft for the README file for the `Remove-EEEU` repository:

---

# Remove-EEEU

Scripts to handle EEEU from the file level in SharePoint Online.

## Overview

This repository contains a collection of PowerShell scripts designed to manage and remove the EEEU (Empty Elements, Empty Uses) from SharePoint Online files. These scripts help in maintaining the cleanliness and efficiency of your SharePoint Online environment by identifying and eliminating unnecessary elements.

## Repository Contents

- **Add-EEEU.ps1**:  Adds the "Everyone Except External Users" (EEEU) group with specified permissions to a file in SharePoint Online.
- **Add-EEEUtoOneDrive-RootandLibrary.ps1**: Script to add Everyone Except External Users (EEEU) permissions to a SharePoint site and its document library.
- **Find-EEEUInSites.ps1**: Scans SharePoint Online sites to identify occurrences of the "Everyone Except External Users" (EEEU) group in file permissions.
- **Find_EEEU_Root_Library.ps1**: Scans OneDrive sites for Everyone Except External Users (EEEU) permissions at root and document library level.
- **Find_and_Remove_EEEU_From_Files_in_OneDrive.ps1**: Find and Removes specific user permissions ("EEEU") from files in a Onerive Site.
- **Remove-EEEUFromFileList.ps1**: Script to remove Everyone Except External Users (EEEU) permissions from files listed in a CSV.
- **Remove-EEEU_from_File.ps1**:
- **Remove-EEEU_from_Files_in_Sites_List.ps1**:

<!-- Add descriptions for each script in the repository -->

## Prerequisites

Before running these scripts, ensure you have the following:

- PowerShell 7.2
- PNP PowerShell Module
- Necessary Graph permissions to execute scripts on SharePoint Online

## Usage

1. Clone the repository to your local machine:
    ```sh
    git clone https://github.com/mikelee1313/Remove-EEEU.git
    cd Remove-EEEU
    ```

2. Open PowerShell with administrative privileges.

3. Run the desired script:
    ```sh
    .\Find-EEEUInSites.ps1
    ```

4. Follow any on-screen instructions provided by the script.

## Contributing

Contributions are welcome! If you have suggestions or improvements, please create a pull request or open an issue.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Author

- [Mike Lee](https://github.com/mikelee1313)

---

Feel free to modify the script descriptions and other sections to better fit the specifics of your repository.
