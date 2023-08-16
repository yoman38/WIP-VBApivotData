## Employee Shift Data Processor and Excel to SQL Command Generator

### Employee Shift Data Processor

The Employee Shift Data Processor is a versatile tool designed to efficiently process and analyze employee shift data along with work statuses. Utilizing VBA (Visual Basic for Applications), this tool seamlessly integrates into an Excel VBA module. The tool's workflow is summarized as follows:

1. **User Input**: Users are prompted to provide essential details, including the target worksheet, the last row, and the column containing 'Brygada' information.

2. **Subroutines**: The tool engages two distinct subroutines, namely `PivotData` and `PivotDataUnique`, utilizing user-provided input.

3. **Empty Cell Check**: After processing, the tool meticulously inspects the created sheets for any empty cells, promptly alerting users if any are found.

### Functions

1. **GetUserInputAndRunSubroutines()**: This central function orchestrates the tool's operations. It captures user input and triggers relevant subroutines as required.

2. **PivotData(ByVal wsNum As Long, ByVal lastRow As Long, ByVal squadCol As String)**: This function generates two new sheets, "WorkersShifts" and "WorkersMonthData," and populates them with data from the specified source worksheet. Data is methodically processed row by row to populate the new sheets with pertinent information.

3. **PivotDataUnique(ByVal wsNum As Long, ByVal lastRow As Long, ByVal squadCol As String)**: Creating a new sheet named "WorkersStatus," this function fills it with unique values from the provided source worksheet. It meticulously ensures that duplicates are excluded, and the sheet exclusively holds distinct worker names and their corresponding details.

4. **CheckForEmptyCells()**: This function scrutinizes the sheets "WorkersShifts," "WorkersMonthData," and "WorkersStatus" for any lurking empty cells. Detecting any, it promptly informs users of their presence.

### How to Use the Tool

1. Click the designated button to launch the tool.
2. Respond to the ensuing dialog boxes by entering the requested inputs.
3. Witness the tool create new worksheets and expertly populate them with meticulously processed data.

**Note**: Prior to employing the tool, ensure that macros are enabled, or adjust your Excel macro security settings to facilitate their execution. Manage these settings in Excel's options under the `Trust Center Settings`.

### Excel to SQL Command Generator (Integrated within Employee Shift Data Processor)

Contained within the Employee Shift Data Processor is the potent "Excel to SQL Command Generator." This module significantly streamlines the conversion of Excel data into SQL commands, effectively generating SQL statements from Excel files that are then stored in text format. The module's mission is to expedite the process of transforming Excel data into SQL queries for seamless database operations. Here's a comprehensive overview of its key features and usage guidelines:

#### Key Features

- Create SQL commands for table creation and data insertion.
- Specify Excel input file and output text file paths.
- Filter rows based on specific keywords.
- Exclude duplicate rows by designated columns.
- Remove rows with empty cells within a defined range.
- Incorporate unique identifiers for rows.
- Define primary keys for the table.
- Set up foreign keys within the table.
- Define NOT NULL constraints for columns.
- Implement indexes and constraints within the table.
- Assign default values to columns.
- Flawlessly handle Latin special characters from various European languages for SQL compatibility.
- Establish a direct connection to SQL Server and execute SQL statements from a `.txt` file.
- Ensure data integrity by purging existing tables before generating new ones.

#### Usage Instructions

**Latest Release: Enhanced Usability**

1. **Installation**: Download the `.xlsm` file.

2. **Run the Main Procedure**:
   - Click the "Run" button.
   - A dialog box will prompt for inputs:
     - Select the Excel file.
     - Designate the output text file for storing generated SQL commands.
     - Optionally, specify keyword-based row filtering (e.g., "yes" or "no").
     - If filtering, provide the keyword for retrieving specific rows.
     - Optionally, skip duplicate entries.
     - Optionally, skip rows with empty cells.
     - Optionally, enable unique ID generation.

3. **VBA Processing**:
   - The VBA code opens the Excel file.
   - Data is extracted based on specified ranges.
   - SQL statements are generated for table creation and data insertion.

4. **Output**:
   - The generated SQL statements are saved in the designated output text file.

5. **Handling Polish Special Characters**:
   - The code adeptly manages Polish special characters to ensure SQL compatibility.

6. **Automated SQL Server Connection**.

**Previous Installation and Usage**

- Installation: Download the provided `.xlsm` file.
- Usage:
   - Open the downloaded `.xlsm` file in Excel.
   - Access the VBA editor by pressing "Alt + F11."
   - Create a new module through "Insert" in the top menu and select "Module."
   - Copy and paste the provided code into the new module.
   - Add a button and assign the macro.
   - Execute the main procedure "GenerateSQL" by clicking "Run" or pressing "F5."
   - Follow the prompts to input essential details, including Excel file path, output text file path, filtering preferences, unique IDs, keys, constraints, indexes, and more.
   - The VBA code adeptly handles Excel data, generating SQL statements as specified.
   - The resultant SQL statements are stored in the output text file.

## Important Considerations

- **Backup**: Prior to code execution, ensure a backup of the Excel file exists. The code may manipulate Excel data during operations.
- **Sheet Selection**: Incorrect active sheet selection may lead to issues. Utilize "RESTRICTION" mode for accuracy. (Note: DELETED NOW)
- **Mixed Data Type Detection**: Version v2 addressed mixed data type detection using NVARCHAR.
- **Clarity and Prompting**: Versions v3 and v4 improved user clarity and prompts during execution.
- **Review Generated SQL**: The code may open the output text file for your review before uploading.

## Changelog

### v3
- Resolved an issue causing merged values between runs when the workbook wasn't closed.
- Enhanced selection range for improved accuracy and reliability.

### v3.11
- Ensured active sheet activation before user range selection.

### v3.35
- Refined range selection process for enhanced usability.

### v3.36
- Automatically opens the generated output text file for review before uploading.
- Rectified output issues and related concerns.
- Expanded options for specifying keys.

### v3.4
- Introduced the option to skip rows containing specific words.
- Enhanced prompts for better user comprehension.
- Enabled entering multiple filter words

.
- Data range is now selected for the filter range by default.
- Now handles all European languages with a Latin alphabet, not just Polish.

### v3.43
- Now handles Latvian, Hungarian, etc. Converted alphabet to their Latin counterpart to ensure functionality in SQL.
- Cancelled the addition of Russian, Greek, etc. These characters weren't correctly displayed on my computer.

### Additional Enhancements
- Included additional prompts for clearer guidance and options.

### Latest Versions:

#### v3.44
- Levenshtein method to find the distance between words: Work in progress, not currently functioning (omitted in the output, possibly an issue in the loop). The goal is to make it work even if users write "Lisyoped" instead of "Linstopad".

#### v3.45
- Allow duplicates of WorkersName in WorkersStatus (if one worker is part of two brygada).
- Added, then deleted, the EmptyCellsCheck function opening: 1) a new worksheet (causing troubles with the active sheet), 2) a new workbook (too long loading), 3) a txt file (faster loading, could be used… but still a bit slower)

#### v3.46
- Finally added support for other Latin characters (Latin2 encoding).
- Other minor improvements such as error handling.

#### v3.47
- Corrected an important error causing the code to stop working.
- Sub CheckForEmptyCells(srcWorkbook As Workbook) now saves a txt file in downloads instead of opening it.
- Added back the option to open it.
- Conflict names added.

These updates have been integrated into the tool to enhance its functionality, usability, and overall user experience.
---

## Further Information

The subsequent section delves into additional intricacies and elucidates each feature through illustrative examples:

---

---

## Additional Details

Before generating an output, the following queries may arise:

- **Omitting Rows without a Specific Keyword**: This function enables row exclusion based on a designated keyword. For instance, with the following dataset:
  ```
  1 / Michael / 123
  2 / Julius / 456
  1 / Michael / 123
  1 / Olga / 789
  3 / Kevin / 789
  ```
  Selecting data range A1:C4 and filtering range A1:A5 with keyword "1" results in:
  ```
  1 / Michael / 123
  1 / Michael / 123
  1 / Olga / 789
  ```
  Choosing data range A1:C4 and filter range A1:A4 with keyword "1" results in:
  ```
  1 / Michael / 123
  1 / Michael / 123
  1 / Olga / 789
  3 / Kevin / 789
  ```
  Similarly, selecting range C1:C5 with keyword "123" yields:
  ```
  1 / Michael / 123
  1 / Michael / 123
  ```

- **Excluding Duplicate Rows Based on Specific Columns**: This feature facilitates the removal of duplicate rows based on designated columns. Using column A:
  ```
  1 / Michael / 123
  2 / Julius / 456
  ```
  And with column C:
  ```
  1 / Michael / 123
  2 / Julius / 456
  1 / Olga / 789
  ```

- **Omitting Rows with Empty Cells in a Specific Range**: This functionality discards rows containing empty cells within a specified range.

- **Inclusion of a Unique ID for Each Row**: Activation of this feature appends an ID column "[Id] [int] IDENTITY(1,1) NOT NULL" for each row in the table.

- **Setting Primary Key for the Table**: This choice empowers the establishment of a primary key for the table. Multiple columns can be selected and set as NOT NULL via the "ALTER TABLE ... ADD PRIMARY KEY ..." statement.

- **Adding a Foreign Key**: Selection of a column, its NOT NULL assignment, reference table, and column creation results in an "ALTER TABLE ... ADD FOREIGN KEY ... REFERENCES ..." statement.

- **Imposing NOT NULL Constraint on Other Columns**: This attribute streamlines the process of setting specific columns as NOT NULL through the "ALTER TABLE ... ALTER COLUMN ... NOT NULL" statement.

- **Inclusion of an Index**: Options include "Non-clustered," "Clustered," or "Unique Non-clustered with Sort Order." Column selection leads to the creation of 'CREATE INDEX' / 'CREATE

 CLUSTERED INDEX' / 'CREATE UNIQUE INDEX … ON … DESC' (or ASC).

- **Addition of a Constraint**: Constraints such as "unique" or "check" can be set using the "ALTER TABLE … ADD CONSTRAINT …" statement.

- **Assignment of Default Value for Any Column**: Utilize this functionality to assign default values to columns via the "ALTER TABLE … ADD CONSTRAINT … DEFAULT … FOR …" statement.

---
