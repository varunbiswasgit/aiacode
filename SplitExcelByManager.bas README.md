# Split Excel by Manager (VBA Macro)

## Overview
`SplitExcelByManager.bas` is a VBA macro for Microsoft Excel. It splits the active worksheet into separate workbooks based on the manager names found in a specified column (default `Manager`). Each new workbook contains the rows for one manager.

## Usage
1. Open your source workbook in Excel.
2. Press `Alt + F11` and import `SplitExcelByManager.bas`.
3. Run `SplitExcelByManager` from the macros list.
   - When prompted, enter the column header containing manager names.
  - Choose the destination folder when the folder picker appears.
   - The macro now filters using the full data range to avoid AutoFilter errors.

## Example
Running the macro with the default `Manager` column creates files like `Alice.xlsx` and `Bob.xlsx` in the folder you selected.

## Task List for SplitExcelByManager.bas
- [x] Allow user to choose a different save location
- [ ] Show progress while creating files
- [ ] Handle protected or filtered worksheets gracefully

## License
This script is released under the [GNU General Public License v3.0](LICENSE).
