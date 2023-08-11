Attribute VB_Name = "Excel2MSSQL"

' Declare a public variable to store the path of the selected Excel file
Public gSelectedExcelFile As String
Sub MainProcess()



End Sub



Function GetUserInputAndRunSubroutines0() As Boolean

    ' Ask the user to select the source file
    Dim srcFile As Variant
    srcFile = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*), *.xls*", _
                                          Title:="Please select the source Excel file")
                                          
    ' Exit if the user didn't choose a file
    If srcFile = False Then
        MsgBox "No file selected. Exiting..."
        GetUserInputAndRunSubroutines0 = False
        Exit Function
    End If

    ' Open the selected workbook
    Dim srcWorkbook As Workbook
    Set srcWorkbook = Workbooks.Open(srcFile)
    gSelectedExcelFile = srcFile

    ' Return True since the file was successfully opened
    GetUserInputAndRunSubroutines0 = True

End Function

Sub RunAllProcesses0()

    ' Call the main procedure from TPtoPivot and check its result
    ' Assuming that there's some function or code that should be here to determine the success/failure of the TPtoPivot process
    
    ' For now, I'll assume it always succeeds, but you should replace the True with the actual check
    If True Then
        
        ' Call the main procedure from ExcelSQL
        GenerateSQL
        
        ' Prompt the user to check if they want to upload the table to Microsoft SQL Server
        Dim userResponse As VbMsgBoxResult
        userResponse = MsgBox("Do you want to upload your table to Microsoft SQL Server?", vbYesNo)
        
        If userResponse = vbYes Then
            ' Call the main procedure from Txt2SQL
            UpdateSQLWithTxtContent
        End If
        
    Else
        ' Do not run other modules
        MsgBox "Processes were cancelled or not completed successfully. Other modules will not run.", vbExclamation
    End If
End Sub


