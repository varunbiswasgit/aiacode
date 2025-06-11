# Split Excel by Manager VBA Macro

This VBA macro splits the active worksheet into separate workbooks based on the value in the **Manager** column (column A by default). For each unique manager, a new workbook is created containing only the rows for that manager.

Each output workbook automatically adjusts column widths so that all data is visible.

## Usage
1. Open the workbook containing the data.
2. Press `Alt+F11` to open the VBA editor and import `SplitExcelByManager.bas`.
3. Run `SplitExcelByManager` from the editor.

The macro will create a workbook for each manager in the same folder as the original file.
