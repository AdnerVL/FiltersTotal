# Filter Quantity Calculator

This project provides a simple batch script (`Filters.bat`) that processes an Excel file to calculate the total quantities of filters by size and outputs the results to a CSV file. It’s designed for users who need to tally filter orders from a spreadsheet without manual effort.

## Purpose

The script reads an Excel file in the same directory, where:
- Column B (starting at B2) lists filter sizes (e.g., "20x20x1").
- Column C (starting at C2) lists quantities.
It groups identical sizes, sums their quantities, and saves the totals to `filter_totals.csv`.

## Requirements

- **Windows**: The script runs on Windows via Command Prompt.
- **PowerShell**: Included with Windows, used to execute the script.
- **Excel File**: An `.xlsx` file with filter data (no headers required).
- **Internet (first run)**: To install the `ImportExcel` PowerShell module if not already present (requires admin rights).

## How to Use

1. **Prepare Your Excel File**:
   - Place your filter data in an `.xlsx` file (e.g., `filters.xlsx`).
   - Ensure sizes are in Column B (B2 onward) and quantities in Column C (C2 onward).
   - Example:

text ```
B      C
20x20x1 5
12x12x1 1
20x20x1 5
12x12x1 1
```

- Save it in the same folder as `Filters.bat` (e.g., `c:\Tools\tests`).

2. **Run the Script**:

- Open Command Prompt.
- Navigate to the script’s folder: `cd c:\Tools\tests`.
- Execute: `Filters.bat`.
- The script finds the first `.xlsx` file in the directory and processes it.

3. **Check the Output**:

- Results are saved to `filter_totals.csv` in the same folder.
- Example output (`filter_totals.csv`):
"Size","TotalQuantity"
"20x20x1","10"
"12x12x1","2"

- Open it in Excel or a text editor to review.

## Script Details

- **File**: `Filters.bat`
- **Process**:
 1. Checks for and installs the `ImportExcel` module if needed.
 2. Finds an `.xlsx` file in the current directory.
 3. Reads data from "Sheet1", starting at B2 (sizes) and C2 (quantities).
 4. Groups sizes, sums quantities, and exports to `filter_totals.csv`.
- **Cleanup**: Deletes the temporary PowerShell script (`temp_script.ps1`) after running.

## Customization

- **Sheet Name**: Edit `-WorksheetName "Sheet1"` in the script if your data is on a different sheet (e.g., `-WorksheetName "Filters"`).
- **Output File**: Change `"filter_totals.csv"` to another name if desired (e.g., `"results.csv"`).
- **File Location**: The script assumes the `.xlsx` file is in the same directory. Move it and adjust paths if needed.

## Troubleshooting

- **No `.xlsx` File**: If no Excel file is found, the script exits with a message.
- **Admin Rights**: First run may require admin privileges to install `ImportExcel`.
- **Errors**: If you see "`) was unexpected`", check for syntax issues or share the output for help.

## License

This is a simple utility script with no formal license—use it freely for your filter-ordering needs!

## Contact

For issues or suggestions, feel free to reach out (or imagine you’re asking an AI assistant who helped build this!)