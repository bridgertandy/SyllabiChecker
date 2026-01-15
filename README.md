This repository holds tools that I made to use for work.


# **Instructions for using the tools.**

### LibraryInstaller.ps1

This script is used to download Python libraries that are necessary to use the Python scripts in this repository.



To run it, right-click on the file and select "Run with PowerShell"



If you get an error about permissions or execution policies, copy, paste, and run the following command and try again.



Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser





### SyllabiChecker.py

This script will check file names for the proper naming convention and it will find schedules. Keep in mind it is CASE-SENSITIVE and VERY SPECIFIC for syllabi. The naming convention it follows is listed in the Syllabi Archive project instructions.



If the file name has "Schedule" or "schedule" anywhere in it, it will catch it. Since these are merged manually later they don't need to follow the naming convention exactly.



3 folders will be created, the NAME ERROR and SCHEDULES are for the tasks above. The FOUND MATCH folder will be used for the MasterListChecker.py script, please leave it alone. Any files that are left over are considered ready to be checked off on the master list.



### MasterListChecker.py

*SyllabiChecker.py MUST be run on the folder before this script can be ran.*



To use this script, run it, then select a folder to sort through, then follow the instructions that appear in the terminal. All files that find a match in the excel spreadsheet will be moved to the FOUND MATCH folder, any files that do not find a match will be left where they are. You can then fix the files that are left over, and run the code again to match the remaining files.



This script also edits the excel file and spreadsheet that you have open to check off the rows that match. It will always check off COLUMN G, so make sure that's the right one. This script cannot uncheck any rows, so clear the cells with X's before you run it (I recommend clearing them one semester at a time in case anything goes wrong).



Keep in mind, the script only matches if the file name EXACTLY matches a row in the spreadsheet. This means the excel spreadsheet may need to be reformatted for the script to run correctly (For example, the term column should be changed to match the "FA20" format instead of "2020FA")



### General tips

If you are going to move or do anything with a large folder (an entire semester folder), do it on the onedrive website, not on your local files. Onedrive syncing between the computer and the website is very slow. Small changes are ok to make on your local files.

