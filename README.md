# Remove-EEEU
Scripts to handle EEEU "Everyone Except for External Users" from the file level in SharePoint Online


**Find-EEEUInSites.ps1** Finds all instances files shared with (EEEU) permissions in SharePoint Online and OneDrive sites.

Output Example:

![image](https://github.com/user-attachments/assets/03c6c701-6682-4198-af46-04d84977822c)

Note: This output CSV file can be directly used with the **Remove-EEEUFromFileList.ps1** to mitigate oversharing using the input list

![image](https://github.com/user-attachments/assets/eb0a6d81-624c-4f3a-9b64-c718e2503b04)


**Remove-EEEUFromFileList.ps1** = Removes EEEU from an input list of files across site. The file list should contain the URL and File Path.

Example input from Find-EEEUInSites.ps1

![image](https://github.com/user-attachments/assets/61962bfa-1d1c-4fb8-8994-a20ca70ce0f9)


**Find-RemoveEEEUfromSites.ps1** This script combines the functionality of "Find-EEEUInSites.ps1" and "Remove-EEEUFromFileList.ps1". 
It first locates all EEEU occurrences using the same method as Find-EEEUInSites.ps1, and then removes the EEEU role from each object as it is found.


![image](https://github.com/user-attachments/assets/b81c8d42-12a7-4652-b9e4-a66a9794e47e)


------------------------------------------------------------

**Disclaimer:** The sample scripts are provided AS IS without warranty of any kind. 
Microsoft further disclaims all implied warranties including, without limitation, 
any implied warranties of merchantability or of fitness for a particular purpose. 
The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
In no event shall Microsoft, its authors, or anyone else involved in the creation, 
production, or delivery of the scripts be liable for any damages whatsoever 
(including, without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use of or inability 
to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

------------------------------------------------------------
