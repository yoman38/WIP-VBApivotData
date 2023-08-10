Attribute VB_Name = "Module1"
Sub GetUserInputAndRunSubroutines()


    ' Get last row from user
    Dim lastRow As Long
    lastRow = InputBox("Please enter the last row number:", "Input needed", "1000")
    
    ' Call both subroutines with the user input
    PivotData lastRow
    PivotDataUnique lastRow

End Sub
Sub PivotData(ByVal lastRow As Long)

    Dim srcSheet As Worksheet
    Dim destSheet As Worksheet, destSheet2 As Worksheet
    Dim i As Long, j As Long, k As Long
    Dim Name As String
    Dim monthYear As String
    Dim day As Variant, data As Variant
    Dim shift As String
    Dim month As Integer, year As String
    Dim monthDict As Object


    ' Dictionary to map Polish month names to month numbers
    Set monthDict = CreateObject("Scripting.Dictionary")
    monthDict("styczeñ") = 1
    monthDict("luty") = 2
    monthDict("marzec") = 3
    monthDict("kwiecieñ") = 4
    monthDict("maj") = 5
    monthDict("czerwiec") = 6
    monthDict("lipiec") = 7
    monthDict("sierpieñ") = 8
    monthDict("wrzesieñ") = 9
    monthDict("paŸdziernik") = 10
    monthDict("listopad") = 11
    monthDict("grudzieñ") = 12

    ' Set source worksheet
    Set srcSheet = ThisWorkbook.Sheets("W5 grafik brygad 2022-2023")

    ' Check if "PivotTable" and "PivotTable2" sheets exist. If not, create them.
    Dim ws As Worksheet
    Dim sheetExists As Boolean, sheetExists2 As Boolean
    sheetExists = False
    sheetExists2 = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "PivotTable" Then
            sheetExists = True
        ElseIf ws.Name = "PivotTable2" Then
            sheetExists2 = True
        End If
    Next ws

    If Not sheetExists Then
        Set destSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        destSheet.Name = "PivotTable"
    Else
        Set destSheet = ThisWorkbook.Sheets("PivotTable")
    End If

    If Not sheetExists2 Then
        Set destSheet2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        destSheet2.Name = "PivotTable2"
    Else
        Set destSheet2 = ThisWorkbook.Sheets("PivotTable2")
    End If

    ' Before you start filling the PivotTable sheets with data, clear the existing contents.
    Application.EnableEvents = False ' disable events
    destSheet.Cells.ClearContents
    destSheet2.Cells.ClearContents
    Application.EnableEvents = True ' enable events again

    destSheet.Cells(1, 1).Value = "Name"
    destSheet.Cells(1, 2).Value = "Date"
    destSheet.Cells(1, 3).Value = "Shift"
    destSheet.Cells(1, 1).EntireRow.Font.Bold = True

    ' Setup headers for PivotTable2
    destSheet2.Cells(1, 1).Value = "Name"
    destSheet2.Cells(1, 2).Value = "Date"
    destSheet2.Cells(1, 3).Value = "Header"
    destSheet2.Cells(1, 4).Value = "Data"
    destSheet2.Cells(1, 1).EntireRow.Font.Bold = True

    Dim destSheetRow As Long, destSheet2Row As Long
    destSheetRow = 2
    destSheet2Row = 2

    For i = 3 To lastRow
        Name = CStr(srcSheet.Cells(i, "G").Value) ' Convert cell content to a string
        monthYear = srcSheet.Cells(i, "H").Value
        
        ' Skip if name is empty, equals 'Nazwisko i imiê', equals '-', or equals '0'
        If Name = "" Or Name = "Nazwisko i imiê" Or Name = "-" Or Name = "0" Then
            GoTo NextRow
        End If

        ' Split monthYear string into month and year
        If InStr(monthYear, "zm.") > 0 Then
            GoTo NextRow
        Else
            month = monthDict(Split(monthYear, " ")(0))
            year = Split(monthYear, " ")(1)
        End If

        ' Check if the month and year for the shift data is the same as the month and year for the day data
        Dim nextMonthYear As String
        nextMonthYear = srcSheet.Cells(i + 1, "H").Value
        Dim nextMonth As Integer, nextYear As String
        If InStr(nextMonthYear, "zm.") > 0 Then
            nextMonth = monthDict(Split(nextMonthYear, " zm. ")(0))
            nextYear = Split(nextMonthYear, " zm. ")(1)
        Else
            GoTo NextRow
        End If
        
        If nextMonth <> month Or nextYear <> year Then
            MsgBox "Error: Month and year for rows " & i & " and " & i + 1 & " do not match."
            GoTo NextRow
        End If

        ' Add data to PivotTable2 for columns AT to BF
        For k = 46 To 58
            data = srcSheet.Cells(i + 1, k).Value
            If Not IsEmpty(data) And Not IsError(data) Then
                destSheet2.Cells(destSheet2Row, 1).Value = Name
                destSheet2.Cells(destSheet2Row, 2).Value = DateSerial(year, month, 1)
                destSheet2.Cells(destSheet2Row, 2).NumberFormat = "yyyy-mm-dd" ' Change date format here
                destSheet2.Cells(destSheet2Row, 3).Value = srcSheet.Cells(2, k).Value
                destSheet2.Cells(destSheet2Row, 4).Value = data
                destSheet2Row = destSheet2Row + 1
            End If
        Next k

        ' Add data to PivotTable for columns I to AS
        For j = 9 To 44
            day = srcSheet.Cells(i, j).Value
            shift = srcSheet.Cells(i + 1, j).Value
            
            ' Only process cells that contain a numeric day value
            If IsNumeric(day) Then
                
                If Not IsEmpty(shift) And Not IsError(shift) Then
                    destSheet.Cells(destSheetRow, 1).Value = Name
                    destSheet.Cells(destSheetRow, 2).Value = DateSerial(year, month, day)
                    destSheet.Cells(destSheetRow, 2).NumberFormat = "yyyy-mm-dd" ' Change date format here
                    destSheet.Cells(destSheetRow, 3).Value = shift
                    destSheetRow = destSheetRow + 1
                End If
            End If
        Next j

NextRow:
    Next i

    ' Autofit columns in the PivotTable sheet
    destSheet.Columns("A:C").EntireColumn.AutoFit
    
    ' Autofit columns in the PivotTable2 sheet
    destSheet2.Columns("A:D").EntireColumn.AutoFit
    
    ' Cleanup
    Set srcSheet = Nothing
    Set destSheet = Nothing
    Set destSheet2 = Nothing
    Set monthDict = Nothing
    
End Sub
Sub PivotDataUnique(ByVal lastRow As Long)
    
    Dim srcSheet As Worksheet
    Dim destSheet3 As Worksheet
    Dim i As Long
    Dim Group As String
    Dim Squad As String
    Dim Abbreviation As String
    Dim Name As String
    Dim dictUnique As Object
    Dim GroupRelevant As Integer
    
    ' Set source worksheet
    Set srcSheet = ThisWorkbook.Sheets("W5 grafik brygad 2022-2023")
    
    ' Check if "PivotTable3" sheet exists. If not, create it.
    Dim ws As Worksheet
    Dim sheetExists3 As Boolean
    sheetExists3 = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "PivotTable3" Then
            sheetExists3 = True
        End If
    Next ws
    
    If Not sheetExists3 Then
        Set destSheet3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        destSheet3.Name = "PivotTable3"
    Else
        Set destSheet3 = ThisWorkbook.Sheets("PivotTable3")
    End If

    ' Clear PivotTable3 sheet before adding new data
    destSheet3.Cells.ClearContents

    GroupRelevant = MsgBox("Is the column before Squad relevant?", vbYesNo)
    
    If GroupRelevant = vbYes Then
        destSheet3.Cells(1, 1).Value = "Group"
        destSheet3.Cells(1, 2).Value = "Squad"
        destSheet3.Cells(1, 3).Value = "Abbreviation"
        destSheet3.Cells(1, 4).Value = "Name"
    Else
        destSheet3.Cells(1, 1).Value = "Squad"
        destSheet3.Cells(1, 2).Value = "Abbreviation"
        destSheet3.Cells(1, 3).Value = "Name"
    End If

    destSheet3.Cells(1, 1).EntireRow.Font.Bold = True
    
    Dim destSheet3Row As Long
    destSheet3Row = 2

    Set dictUnique = CreateObject("Scripting.Dictionary")

    For i = 3 To lastRow
        Group = CStr(srcSheet.Cells(i, "D").Value) ' Convert cell content to a string
        Squad = CStr(srcSheet.Cells(i, "E").Value) ' Convert cell content to a string
        Abbreviation = CStr(srcSheet.Cells(i, "F").Value) ' Convert cell content to a string
        Name = CStr(srcSheet.Cells(i, "G").Value) ' Convert cell content to a string

        ' Skip if name is empty, equals 'Nazwisko i imiê', equals '-', or equals '0'
        If Name = "" Or Name = "Nazwisko i imiê" Or Name = "-" Or Name = "0" Then
            GoTo NextRowUnique
        End If

        ' Only add unique values to PivotTable3
        If Not dictUnique.exists(Name) Then
            dictUnique(Name) = ""

            If GroupRelevant = vbYes Then
                destSheet3.Cells(destSheet3Row, 1).Value = Group
                destSheet3.Cells(destSheet3Row, 2).Value = Squad
                destSheet3.Cells(destSheet3Row, 3).Value = Abbreviation
                destSheet3.Cells(destSheet3Row, 4).Value = Name
            Else
                destSheet3.Cells(destSheet3Row, 1).Value = Squad
                destSheet3.Cells(destSheet3Row, 2).Value = Abbreviation
                destSheet3.Cells(destSheet3Row, 3).Value = Name
            End If
            
            destSheet3Row = destSheet3Row + 1
        End If

NextRowUnique:
    Next i

    ' Autofit columns in the PivotTable3 sheet
    destSheet3.Columns("A:D").EntireColumn.AutoFit
    
    ' Cleanup
    Set srcSheet = Nothing
    Set destSheet3 = Nothing
    Set dictUnique = Nothing
End Sub

