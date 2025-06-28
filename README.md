Overview
COMPILE_EXCEL is a powerful VBA macro designed to combine data from multiple Excel and CSV files into a single worksheet. It supports combining all sheets or a specific sheet across files, handles varied column headers, skips duplicates (optionally), and logs errors and issues encountered during the merge process.

Features
Combine Multiple Files: Merges data from .xls, .xlsx, .xlsm, and .csv files in a chosen folder.
Flexible Sheet Selection: Option to combine all sheets or just a specific sheet name.
Header Handling: Dynamically gathers all unique headers across files and allows custom column order.
Duplicate Removal: Optionally skips duplicate rows based on all columns.
Error Logging: Detailed error log sheet tracks files/sheets that could not be opened or processed.
Safe Naming: Automatically renames existing output/error log sheets to prevent overwriting.
Status Updates: Application status bar and summary messages keep you informed.
How to Use
Open the Macro-Enabled Workbook:
Place the VBA code in a module in a macro-enabled workbook (.xlsm).

Run the Macro:
Use the CombineExcelFilesWithAllEnhancements macro to start. (You may need to enable macros.)

Select Folder:
The macro will prompt you to select a folder containing the Excel/CSV files to combine.

Choose Sheet(s):
You can choose to combine all sheets or specify a particular sheet name.

Set Column Order (Optional):
Optionally specify a custom column order for the combined output.

Handle Duplicates (Optional):
Choose whether to skip duplicate rows (based on all columns).

Review Results:

The combined data will be output to a new CombinedData worksheet.
An error log is generated in a separate sheet (e.g., CombineErrors).
Existing output/error log sheets are backed up with unique names.
Key Macro Functions
SheetExists(wb, nm): Checks if a worksheet exists in a workbook.
GetUniqueSheetName(wb, baseName): Finds a valid, unique sheet name (max 31 chars).
CombineExcelFilesWithAllEnhancements: Main subroutine that manages the entire combine process.
Notes
Supports .xls, .xlsx, .xlsm, and .csv file types.
Skips the workbook containing the macro itself.
Handles invalid sheet names and Excel sheet name restrictions.
Error log helps troubleshoot issues such as missing files or sheets.
Uses late binding for FileSystemObject (no extra references needed).
Requirements
Microsoft Excel (with macro support enabled)
Basic knowledge of how to run VBA macros
Disclaimer
This macro will modify the workbook where it is run. Ensure you have backups of your data before running.

Feel free to copy and adapt this README for your project! If you want the README in a specific format or with additional sections, let me know.

CREATE REDME FILE
Here’s a README file for your project based on the provided VBA (.bas) code:

COMPILE_EXCEL
Overview
COMPILE_EXCEL is a VBA macro designed to automate the process of combining data from multiple Excel and CSV files in a folder. It offers advanced header handling, duplicate removal, error logging, and flexible sheet selection—all within Microsoft Excel.

Features
Merge Multiple Files: Combines data from .xls, .xlsx, .xlsm, and .csv files in a selected folder.
Sheet Selection: Option to combine all sheets or only a specific sheet name from each file.
Header Management: Detects all unique headers, supports custom column order, and aligns columns automatically.
Duplicate Handling: Optionally skips duplicate rows based on all columns.
Error Logging: Generates a detailed error log sheet for troubleshooting.
Safe Sheet Naming: Prevents overwriting existing sheets by auto-renaming.
User Prompts: Interactive dialogs guide you through every step.
How to Use
Open the Macro-Enabled Workbook:
Copy the provided VBA code into a module in an .xlsm file.

Run the Macro:
Execute CombineExcelFilesWithAllEnhancements from Excel (press Alt+F8 and select the macro).

Follow Prompts:

Select a folder containing your Excel/CSV files.
Choose whether to combine all sheets or just a specific sheet.
Optionally set a custom column order.
Decide whether to skip duplicate rows.
Check Results:

Combined data appears in a new CombinedData sheet.
Errors and issues are logged in a separate sheet (e.g., CombineErrors).
Key Functions
SheetExists: Checks if a worksheet exists in a workbook.
GetUniqueSheetName: Generates a valid, unique sheet name.
CombineExcelFilesWithAllEnhancements: The main subroutine that processes, merges, and logs data.
Requirements
Microsoft Excel with macros enabled.
Basic familiarity with running VBA macros.
Notes
The macro will not process the workbook containing the macro itself.
Existing CombinedData and error log sheets are automatically backed up.
Error log helps identify files or sheets that couldn't be opened or processed.
Disclaimer
Back up your data before running macros. This tool modifies the workbook where it is executed.
