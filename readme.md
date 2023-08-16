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
---

---

## Code breakdown


1. **Modules Breakdown:**
   The code is organized into several modules, each containing subroutines and functions that perform specific tasks. These modules include:
   
   - **TP TO PIVOT:** This module handles tasks related to processing data, generating pivot tables, and checking for empty cells.
   - **EXCELSQL:** This module deals with generating SQL statements and interacting with Excel workbooks.
   - **Txt2SQL:** This module focuses on establishing connections to a SQL Server database, executing SQL queries, and managing tables.
   - **MASTERS MODULES:** These modules seem to orchestrate the execution of other modules based on user input and manage the flow of the application.

2. **TP TO PIVOT Module:**
   - `GetUserInputAndRunSubroutines`: This subroutine collects user input (source Excel file, worksheet, etc.), then activates two other subroutines: `PivotData` and `PivotDataUnique`.
   - `PivotData`: Processes data from a worksheet, creating pivot tables and mapping month names to numbers.
   - `PivotDataUnique`: Processes data to populate a worksheet with unique worker information.
   - `CheckForEmptyCells`: Scans worksheets for empty rows or cells and generates a report.
   - `LevenshteinDistance`: Calculates Levenshtein distance between two strings.
   - `GetClosestMonth`: Finds the closest month name from a dictionary based on Levenshtein distance.

3. **EXCELSQL Module:**
   This module provides functions and subroutines to interact with Excel workbooks and generate SQL statements.
   
   The functions and subroutines include:
   - Excel workbook manipulation functions (`OpenWorkbook`, `SelectRange`, `SelectWorksheet`, etc.).
   - SQL statement generation functions (`GenerateCreateTable`, `GenerateInsertStatements`, etc.).
   - Input and user interaction functions (`GetUserInput`, `GetUserResponse`, etc.).
   - Data type guessing function (`GuessDataType`).
   - Functions to handle special characters, file writing, and more.

4. **Txt2SQL Module:**
   This module focuses on managing SQL Server connections and executing SQL queries.
   
   The functions and subroutines include:
   - `ConnectToSQL`: Establishes a connection to a SQL Server database.
   - `GetTableNames`: Retrieves non-system table names from the connected database.
   - `DeleteTableIfExists`: Deletes a specified table from the database if it exists.
   - `UpdateSQLWithTxtContent`: Establishes a connection and executes SQL queries from a text file.
   - File reading and handling functions (`GetQueryFromTxt`, `WriteToFile`).
   
5. **MASTERS MODULES:**
   These modules provide a higher-level orchestration of the application flow based on user input and execution results.
   
   The modules include:
   - **Excel2SQLconverter**: Invokes the SQL generation process, asks for user confirmation, and manages the table transfer process.
   - **TP2SQLconverter**: Similar to the above, but for the "TP TO PIVOT" functionalities.

6. **Application Flow:**
   - The user starts by running one of the "Masters Modules" (e.g., `Excel2SQLconverter` or `TP2SQLconverter`).
   - The module collects necessary input from the user or runs specific subroutines.
   - If the user confirms, SQL statements are generated and executed on a SQL Server database using the `Txt2SQL` module.
   - Various Excel workbook manipulations, pivot table creation, and data processing occur using the `TP TO PIVOT` and `EXCELSQL` modules.
   
7. **Overall Purpose:**
   The code appears to offer functionalities to transform data in Excel sheets into SQL Server databases, generate pivot tables, handle user input and interactions, and execute SQL queries based on user decisions.


---
## Deep explanation

# module TP TO PIVOT

1. **GetUserInputAndRunSubroutines:**
   - Collects user input for source Excel file, worksheet, last table row, and column.
   - Activates chosen worksheet and calls "PivotData" and "PivotDataUnique" subroutines.
   - Returns True if successful, False if input is invalid or operation is canceled.

2. **PivotData:**
   - Processes data from a source worksheet, creating "WorkersShifts" and "WorkersMonthData" PivotTable sheets.
   - Extracts worker names, shift data, and month details.
   - Utilizes a dictionary to map Polish month names to numbers, performs operations, and formats data.

3. **PivotDataUnique:**
   - Processes source data, populating "WorkersStatus" worksheet with unique worker information.
   - Filters duplicates based on name and abbreviation, adds unique info, adjusts columns, and performs cleanup.

4. **CheckForEmptyCells:**
   - Scans specified sheets for empty rows or cells.
   - Generates a report with row content for any issues found.
   - Saves report to a text file and offers to open it using Notepad if relevant rows are detected.

5. **LevenshteinDistance:**
   - Calculates Levenshtein distance between two input strings.
   - Measures minimum edits needed to transform one string into another.
   - Returns the resulting distance as the output.

6. **GetClosestMonth:**
   - Takes a month name and a dictionary of month names.
   - Calculates Levenshtein distance between input month and dictionary entries.
   - Returns closest matched month if distance is below 5, otherwise an empty string.

## module EXCELSQL

1. **GenerateSQL Subroutine:**
   - Initializes variables and settings for SQL query generation.
   - Opens an Excel workbook, selects a worksheet, and defines data ranges.
   - Enables user interaction for filtering, exclusion, duplicate checks, and empty cell handling.
   - Generates SQL statements for table creation, data insertion, and additional details.
   - Writes SQL output to a file, closes the workbook, and displays a success message.
   - Opens the output file for review if available.

2. **GetUserInput Function:**
   - Displays an input box to gather user input based on a prompt.
   - Returns the entered string or a default value if provided.

3. **GetUserResponse Function:**
   - Displays a yes/no input box to retrieve user responses.
   - Returns the user's response as a string.

4. **SelectFile Function:**
   - Displays a file selection dialog and returns the selected file's path.

5. **SelectFolder Function:**
   - Displays a folder selection dialog and returns the selected folder's path.

6. **OpenWorkbook Function:**
   - Opens an Excel workbook at a specified file path.
   - Returns the opened workbook object.

7. **SelectRange Function:**
   - Activates a specified worksheet and prompts user to select a range.
   - Handles errors, ensures correct worksheet selection, and returns the chosen range.

8. **GenerateCreateTable Function:**
   - Creates a SQL CREATE TABLE statement based on headers, data ranges, and output file name.
   - Determines column data types by analyzing data.
   - Constructs SQL statement for table creation.

9. **GenerateInsertStatements Function:**
   - Generates SQL INSERT statements for data in specified range.
   - Includes customizable filtering and conditions for row inclusion.
   - Handles duplicates, empty cells, and generates SQL for qualified rows.

10. **WriteToFile Function:**
    - Writes provided text content to a specified file.
    - Uses the Scripting.FileSystemObject for file handling.

11. **ReplaceSpecialCharacters Function:**
    - Standardizes special characters from various languages into ASCII equivalents.
    - Enhances compatibility for processing or display.

12. **GetBaseName Function:**
    - Extracts base name (filename without extension) from a file path.

13. **SelectWorksheet Function:**
    - Allows user to choose a worksheet within a workbook by index.
    - Validates input and returns selected worksheet.

14. **GuessDataType Function:**
    - Determines and returns guessed data type based on input value characteristics.

15. **GetUserNumber Function:**
    - Prompts user to enter a numeric value using an input box.
    - Validates input and returns a validated number or -1 if canceled.

16. **GetAdditionalTableDetails Function:**
    - Generates SQL statements for modifying a table's structure in a database.
    - Handles Primary Key, Foreign Key, NOT NULL constraints, indexes, constraints, and default values.
    - Interacts with user through prompts and constructs corresponding SQL statements.
    - Returns generated SQL statements for database changes.

# Module Txt2SQL

1. **Function ConnectToSQL(serverName As String, dbName As String) As Boolean:**
   - Establishes a connection to a SQL Server database using server and database names.
   - Utilizes ActiveX Data Objects (ADO) to manage the connection and error handling.
   - Creates ADODB connection and command objects.
   - Sets connection string with user input parameters.
   - Attempts to open the connection and associates the command object.
   - Returns True for successful connection; displays error message and returns False on error.

2. **Function GetTableNames() As String:**
   - Retrieves non-system table names from connected SQL Server database.
   - Uses ADO recordset to access schema information about tables.
   - Excludes system tables, special schemas, and specific prefixes.
   - Accumulates table names in "tableList" variable, separated by line breaks.
   - Returns list of non-system table names as a string.

3. **Subroutine DeleteTableIfExists(tableName As String):**
   - Deletes specified table from the database if it exists.
   - Suppresses errors with "On Error Resume Next".
   - Sets command text to "DROP TABLE" followed by provided table name.
   - Executes command to delete the table.
   - Resets error handling with "On Error GoTo 0" afterward.

4. **Subroutine UpdateSQLWithTxtContent():**
   - Establishes SQL Server connection using user-input server and database names.
   - Displays list of tables using "GetTableNames" function.
   - Retrieves SQL query from text file specified by "outputFilePath".
   - Deletes table if it exists, and drops primary and foreign key constraints.
   - Executes retrieved SQL query using "cmd.Execute".
   - Displays success message if query execution is successful.
   - Cleans up by releasing memory resources for "cmd" and "conn" objects.

5. **Function GetQueryFromTxt(filePath As String) As String:**
   - Reads content of text file specified by "filePath".
   - Uses file system object to handle file operations.
   - Opens file for reading, reads content into "fileContent" variable.
   - Closes file and returns read content as a string.
   - Designed to retrieve SQL queries or other text-based data from files.

## MASTERS MODULES

# Excel2SQLconverter

1. **Excel File Selection:**
   - Asks the user to select an Excel file as the source.
   - If no file is selected, displays a message and exits the subroutine.

2. **Workbook Opening:**
   - Opens the selected Excel workbook.

3. **Main Procedure Invocation (GenerateSQL):**
   - Calls the main procedure "GenerateSQL" from an external module (presumably named ExcelSQL).

4. **User Confirmation:**
   - Asks the user if they want to transfer a table to Microsoft SQL Server.
   - If the user responds affirmatively, proceeds with the table transfer process.

5. **Main Procedure Invocation (UpdateSQLWithTxtContent):**
   - Calls the main procedure "UpdateSQLWithTxtContent" from an external module (presumably named Txt2SQL) if the user wants to transfer the table.

6. **Cancellation Notice:**
   - Displays a message notifying the user that other modules will not run due to cancellations or incomplete processes.


# TP2SQLconverter

1. **Main Procedure Invocation and Validation (GetUserInputAndRunSubroutines):**
   - Calls the main procedure "GetUserInputAndRunSubroutines" to perform user input and execute other subroutines.
   - If the result of "GetUserInputAndRunSubroutines" is True (indicating successful execution), proceeds to the next steps.

2. **Main Procedure Invocation (GenerateSQL):**
   - Calls the main procedure "GenerateSQL" from an external module (presumably named ExcelSQL).

3. **User Confirmation:**
   - Asks the user if they want to transfer a table to Microsoft SQL Server.
   - If the user responds affirmatively, proceeds with the table transfer process.

4. **Main Procedure Invocation (UpdateSQLWithTxtContent):**
   - Calls the main procedure "UpdateSQLWithTxtContent" from an external module (presumably named Txt2SQL) if the user wants to transfer the table.

5. **Cancellation Notice:**
   - If "GetUserInputAndRunSubroutines" returns False (indicating cancellation or unsuccessful execution), displays a message notifying the user that other modules will not run.
