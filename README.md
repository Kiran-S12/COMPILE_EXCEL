# COMPILE_EXCEL

## Overview

**COMPILE_EXCEL** is a powerful VBA macro designed to automate the process of combining data from multiple Excel and CSV files into a single worksheet. It supports merging all sheets or a specific sheet across files, handles varied column headers, can skip duplicate rows (optional), and logs errors and issues during the merge process.

---

## Features

- **Combine Multiple Files:** Merge data from `.xls`, `.xlsx`, `.xlsm`, and `.csv` files within a selected folder.
- **Flexible Sheet Selection:** Option to combine all sheets or just a specific sheet name across files.
- **Header Handling:** Dynamically gathers all unique headers and allows custom column order in the output.
- **Duplicate Removal:** Optionally skips duplicate rows based on all columns.
- **Error Logging:** Generates a detailed error log sheet for files/sheets that could not be processed.
- **Safe Naming:** Automatically renames existing output and error log sheets to prevent overwriting.
- **Status Updates:** Application status bar and summary messages keep you informed of progress.

---

## How to Use

1. **Open the Macro-Enabled Workbook:**
   - Place the VBA code in a module in a macro-enabled workbook (`.xlsm`).

2. **Run the Macro:**
   - Execute the `CombineExcelFilesWithAllEnhancements` macro (you may need to enable macros).

3. **Select Folder:**
   - The macro will prompt you to select a folder containing the Excel/CSV files to combine.

4. **Choose Sheet(s):**
   - Choose to combine all sheets or specify a particular sheet name.

5. **Set Column Order (Optional):**
   - Optionally specify a custom column order for the combined output.

6. **Handle Duplicates (Optional):**
   - Choose whether to skip duplicate rows (based on all columns).

7. **Review Results:**
   - Combined data appears in a new `CombinedData` worksheet.
   - An error log is generated in a separate sheet (e.g., `CombineErrors`).
   - Existing output or error log sheets are automatically backed up with unique names.

---

## Key Macro Functions

- `SheetExists(wb, nm)`: Checks if a worksheet exists in a workbook.
- `GetUniqueSheetName(wb, baseName)`: Finds a valid, unique sheet name (max 31 characters).
- `CombineExcelFilesWithAllEnhancements`: Main subroutine that manages the full combine process.

---

## Requirements

- Microsoft Excel (with macro support enabled)
- Basic knowledge of how to run VBA macros

---

## Notes

- Supports `.xls`, `.xlsx`, `.xlsm`, and `.csv` file types.
- Skips the workbook containing the macro itself.
- Handles invalid sheet names and Excel sheet name restrictions.
- Error log helps identify files or sheets that couldn’t be opened or processed.
- Uses late binding for FileSystemObject (no extra references required).

---

## Disclaimer

This macro will modify the workbook where it is run. **Make sure you have backups of your data before running.**
