# Remove-EEEU
Scripts to handle EEEU "Everyone Except for External Users" from the file level in SharePoint Online


**Find-EEEUInSites.ps1** Finds all instances files shared with (EEEU) permissions in SharePoint Online and OneDrive sites.

Output Example:

![image](https://github.com/user-attachments/assets/03c6c701-6682-4198-af46-04d84977822c)

Note: This output CSV file and be directly used with the **Remove-EEEUFromFileList.ps1** to mitigate oversharing using the input list

**Remove-EEEUFromFileList.ps1** = Removes EEEU from an input list of files across site. The file list should contain the URL and File Path.

Example:

![image](https://github.com/user-attachments/assets/db97140e-282c-45b4-b062-accf760206d2)

Example output:

![image](https://github.com/user-attachments/assets/38aa81c6-d03f-4e41-bcc7-8c97785dc5ea)



**Add-EEEU.ps1** = Script to add EEEU to files for testing

**Add-EEEUtoOneDrive-RootandLibrary.ps1** = Add EEEU to a OneDrive Root Site and Document LIbrary for Testing **MC1013464**

**Find_EEEU_Root_Library.ps1** = This script will scan all OneDrive sites for "Everyone Except External Users" (EEEU) permissions at root and document library level.
This script helps you prepare for **MC1013464** -(Updated) We will remove the EEEU sharing permission from root web and default document library in OneDrive’s.

Example Output:

![image](https://github.com/user-attachments/assets/f2530de2-0194-4857-9088-156b13806646)


![image](https://github.com/user-attachments/assets/18ca2dd1-4108-4d94-beae-cbf7d006d8d8)


**Remove-EEEU.ps1** = Script to remove EEEU from a single file


**Remove_EEEU_From_Files.ps1** =  Removes ("EEEU") from all files including subfolders in a SharePoint Online document library.


**Remove-EEEUFromFilesinSites.ps1**  = This script will take a list of URLs and remove EEEU from all files from all listed sites from within the input file.

Example of input file:

![image](https://github.com/user-attachments/assets/2d01a23b-5896-4c26-ba29-dc1421edb305)

Example of output file:

![image](https://github.com/user-attachments/assets/7a830349-8427-4233-b873-52b5e495ff3d)


**Disclaimer:** The sample scripts are provided AS IS without warranty of any kind. 
Microsoft further disclaims all implied warranties including, without limitation, 
any implied warranties of merchantability or of fitness for a particular purpose. 
The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
In no event shall Microsoft, its authors, or anyone else involved in the creation, 
production, or delivery of the scripts be liable for any damages whatsoever 
(including, without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use of or inability 
to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.
