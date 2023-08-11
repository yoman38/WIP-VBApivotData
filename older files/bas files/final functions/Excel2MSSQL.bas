Attribute VB_Name = "Excel2MSSQL"
Sub RunExcelSQL()

    Dim inputFilePath As String
    inputFilePath = SelectFile
    If inputFilePath = "" Then
        MsgBox "No file selected. Exiting subroutine.", vbExclamation
        Exit Sub
    End If

    ' Call the main procedure from ExcelSQL with the file path as an argument
    GenerateSQL inputFilePath
        
        ' Prompt the user to check if they want to upload the table to Microsoft SQL Server
        Dim userResponse As VbMsgBoxResult
        userResponse = MsgBox("Do you want to upload your table to Microsoft SQL Server?", vbYesNo)
        
        If userResponse = vbYes Then
            ' Call the main procedure from Txt2SQL
            UpdateSQLWithTxtContent
        End If
        
        ' Do not run other modules
        MsgBox "Processes were cancelled or not completed successfully. Other modules will not run.", vbExclamation

End Sub


Function SelectFile() As String
    Dim fd As Object

    Set fd = Application.FileDialog(3)

    With fd
        .Title = "Select an Excel File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        If .Show = -1 Then
            SelectFile = .SelectedItems(1)
        Else
            SelectFile = ""
        End If
    End With

    Set fd = Nothing
End Function

