# Employee Shift Data Processor

This tool processes the employee shift data to analyze shifts and work status. It is implemented in VBA (Visual Basic for Applications), and to use it, you will need to insert the code into an Excel VBA module.

Here is the general flow of the tool:

1. It gets input from the user to specify the worksheet, the last row, and the column letter containing 'Brygada'.
2. The tool calls two subroutines, `PivotData` and `PivotDataUnique`, with the user input.
3. It checks for empty cells in the created sheets.

## Functions

### 1. GetUserInputAndRunSubroutines()

This is the main function that drives the program. It gets input from the user and calls other subroutines.

### 2. PivotData(ByVal wsNum As Long, ByVal lastRow As Long, ByVal squadCol As String)

This function creates two new sheets "WorkersShifts" and "WorkersMonthData" and populates them with data based on the source worksheet, specified by the user. It processes the data row by row and fills the new sheets with relevant data. 

### 3. PivotDataUnique(ByVal wsNum As Long, ByVal lastRow As Long, ByVal squadCol As String)

This function creates a new sheet "WorkersStatus" and populates it with unique values from the specified source worksheet. It avoids duplicates and fills the sheet with unique worker names and their respective information.

### 4. CheckForEmptyCells()

This function checks the created sheets "WorkersShifts", "WorkersMonthData", and "WorkersStatus" for empty cells. If it finds any, it warns the user.

## How to use the tool

1. Open your Excel workbook.
2. Press `Alt + F11` to open the VBA editor.
3. In the VBA editor, insert a new module (Menu -> Insert -> Module).
4. Copy the provided VBA code and paste it into the module.
5. Close the VBA editor.
6. Run the `GetUserInputAndRunSubroutines` procedure (you can press `Alt + F8`, select `GetUserInputAndRunSubroutines`, and press `Run`).
7. Enter the requested inputs in the dialog boxes that pop up.
8. The tool will create new worksheets and fill them with the processed data.

**Note:** Please make sure to enable macros or set your macro security level to a setting that allows them to run. You can adjust these settings in Excel's options under the `Trust Center Settings`.

## Disclaimer

This tool is provided as-is, and you should use it at your own risk. Always make a backup of your Excel files before running any macros or scripts.