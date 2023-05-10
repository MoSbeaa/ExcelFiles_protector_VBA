# ExcelFiles_protector_VBA
This VBA code enables a user to apply different passwords to Excel files located within a folder and its sub-folders. So, each sub-folders files have different password

# Overall, 
this VBA code consists of four subroutines: StartProtecting, getData, gettingCells, and addPassword. 
The StartProtecting subroutine calls the other three subroutines in sequence. 
The getData subroutine prompts the user to select a folder and lists the names of its sub-folders in column A of the active worksheet. 
The gettingCells subroutine prompts the user to select an Excel file and copies data from columns A and B of that file into columns B and C of the active worksheet. 
The makeLoop subroutine loops through each row in a table defined by the used range of cells starting at cell A1 and calls the addPassword subroutine for each row with arguments from columns A and C. 
The addPassword subroutine adds a password to Excel files in a specified folder and its sub-folders.

# How to start
> First you should state all your sub-folders names and the password for each one in "Password List Example" file
> Second you will open "Protect_excel_files_Recursion.xlsm" or "Protect_excel_files_No_Recursion.xlsm" and you will follow the instructions


# Without Recursion
the file "Protect_excel_files_No_Recursion.xlsm" doesn't use recursion and it can go 8 level depth in the sub-folders


# With Recursion
the file "Protect_excel_files_Recursion.xlsm" use recursion and it can go unlimited level depth in the sub-folders

# Differences between with Recursion and without
is that with recursion you can go to unlimited depth of sub-folders but it will take more time to excite 
While without it go only 8 level depth(you can increase them by insert more loops) but it faster when executing

