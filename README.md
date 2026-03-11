Created by Bridger Tandy 2/13/2026

# **Instructions for using the tools.**



### General Instructions

* For any of these to work, your Onedrive has to be synced to your computer. The scripts cannot access the files through the Onedrive website.
* If you are going to move or copy a large folder (an entire semester folder), do it on the Onedrive website, not on your local files. Onedrive syncing between the computer and the website is very slow.
* You will need to download Python for the .py scripts, you can find the standard Python install manager at https://www.python.org/downloads/
* It is easier/faster (not required) to run these scripts in a dedicated Windows terminal, if you know how to use shell commands. Listed below are the only ones you need to know.
* Change directory (cd) to wherever you have these scripts saved. The easiest way to get the path is to find your folder where they are saved, right click, and select copy path.
* To run PowerShell scripts, the command is ./(name of script).ps1
* To run Python scripts (with Python and the libraries already downloaded), the command is py ./(name of script).py



### WordToPdf.ps1

This is used to convert all word documents in a specific folder to PDFs. It is very useful and easy to use. Run the script using the instructions below and select the folder you want to convert.



To run it, right-click on the file and select "Run with PowerShell"



If you get an error about permissions or execution policies, copy, paste, and run the following command in a terminal and try again.



Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser



### LibraryInstaller.ps1

This script is used to download Python libraries that are necessary to use the Python scripts in this folder. Run it the same as the WordToPdf script.



### SyllabiProcesser.py

This is an updated version that just combines the two scripts listed below. Running this will do the same thing as running the other two in order. **USE THIS ONE.** Just open the script by double clicking and it will run.



### SyllabiChecker.py

This script will check file names for the proper naming convention and it will find schedules. Keep in mind it is CASE-SENSITIVE and VERY SPECIFIC for syllabi. The naming convention it follows is listed in the Syllabi Archive project instructions.



If the file name has "Schedule" or "schedule" anywhere in it, it will catch it. Since these are merged manually later they don't need to follow the naming convention exactly.



3 folders will be created, the NAME ERROR and SCHEDULES are for the tasks above. The FOUND MATCH folder will be used for the MasterListChecker.py script, please leave it alone. Any files that are left over are considered ready to be checked off on the master list.



### MasterListChecker.py

*SyllabiChecker.py MUST be run on the folder before this script can be ran.*



To use this script, run it, then select a folder to sort through, then follow the instructions that appear in the terminal. All files that find a match in the excel spreadsheet will be moved to the FOUND MATCH folder, any files that do not find a match will be left where they are. You can then fix the files that are left over, and run the code again to match the remaining files.



This script also edits the excel file and spreadsheet that you have open to check off the rows that match. It will always check off COLUMN G, so make sure that's the right one. This script cannot uncheck any rows, so clear the cells with X's before you run it (I recommend clearing them one semester at a time in case anything goes wrong).



Keep in mind, the script only matches if the file name EXACTLY matches a row in the spreadsheet. This means the excel spreadsheet may need to be reformatted for the script to run correctly (For example, the term column should be changed to match the "FA20" format instead of "2020FA")



### .osts scripts

These are used in Excel under the "Automate" tab. As long as they are in your cloud storage somewhere, Excel should be able to run them. (If you are signed in with the same account)

