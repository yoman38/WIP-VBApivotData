Attribute VB_Name = "Mastersub"
Sub RunAllProcesses()
        
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
