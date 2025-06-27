# Handle EEEU in SharePoint Online
Scripts to handle EEEU "Everyone Except for External Users" from the file level in SharePoint Online

------------------------------------------------------------

**Find-EEEUInSites.ps1** Finds all instances files shared with (EEEU) permissions in SharePoint Online and OneDrive sites.

Output Example:

![image](https://github.com/user-attachments/assets/325d9d06-0580-4353-bc05-6374be00b9af)

![image](https://github.com/user-attachments/assets/9f7cf1b4-0ce1-4ae2-9a47-9de96e8d61a7)

Note: This output CSV file can be directly used with the **Remove-EEEUFromFileList.ps1** to mitigate oversharing using the input list

![image](https://github.com/user-attachments/assets/eb0a6d81-624c-4f3a-9b64-c718e2503b04)

------------------------------------------------------------

**Remove-EEEUFromFileList.ps1** = Removes EEEU from an input list of files across site. The file list should contain the URL, ItemURL and ItemType.

Example input from Find-EEEUInSites.ps1

![image](https://github.com/user-attachments/assets/7633e87b-c304-42fd-afae-86eb9f2eea54)


![image](https://github.com/user-attachments/assets/354ba33c-ffcc-4109-a0d0-24d13eff16b8)


------------------------------------------------------------

**Find-RemoveEEEUfromSites.ps1** This script combines the functionality of "Find-EEEUInSites.ps1" and "Remove-EEEUFromFileList.ps1". 
It first locates all EEEU occurrences using the same method as Find-EEEUInSites.ps1, and then removes the EEEU role from each object as it is found.


![image](https://github.com/user-attachments/assets/b81c8d42-12a7-4652-b9e4-a66a9794e47e)
