Attribute VB_Name = "Module1"

Sub GetUserInputAndRunSubroutines()

    ' Get the name of the source worksheet from user
    Dim ws As Worksheet
    Dim wsNames() As String
    Dim wsNum As Long
    Dim i As Long
    ReDim wsNames(1 To ThisWorkbook.Sheets.Count)
    For Each ws In ThisWorkbook.Sheets
        i = i + 1
        wsNames(i) = i & ". " & ws.Name
    Next ws
    wsNum = InputBox("Please enter the number of the worksheet you want to use as source:" & vbNewLine & _
                      Join(wsNames, vbNewLine), "Input needed", "1")
    
    ' Get last row from user
    Dim lastRow As Long
    lastRow = InputBox("Please enter the last row number:", "Input needed", "100")
    
    ' Get column of "squad" from user
    Dim squadCol As String
    squadCol = InputBox("Please enter the column letter containing 'Brygada':", "Input needed", "E")
    
    ' Call both subroutines with the user input
    PivotData wsNum, lastRow, squadCol
    PivotDataUnique wsNum, lastRow, squadCol

    'check for empty cells
    CheckForEmptyCells
    
End Sub

Sub PivotData(ByVal wsNum As Long, ByVal lastRow As Long, ByVal squadCol As String)

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
    Set srcSheet = ThisWorkbook.Sheets(wsNum)

    ' Check if "WorkersShifts" and "WorkersMonthData" sheets exist. If not, create them.
    Dim ws As Worksheet
    Dim sheetExists As Boolean, sheetExists2 As Boolean
    sheetExists = False
    sheetExists2 = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "WorkersShifts" Then
            sheetExists = True
        ElseIf ws.Name = "WorkersMonthData" Then
            sheetExists2 = True
        End If
    Next ws

    If Not sheetExists Then
        Set destSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        destSheet.Name = "WorkersShifts"
    Else
        Set destSheet = ThisWorkbook.Sheets("WorkersShifts")
    End If

    If Not sheetExists2 Then
        Set destSheet2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        destSheet2.Name = "WorkersMonthData"
    Else
        Set destSheet2 = ThisWorkbook.Sheets("WorkersMonthData")
    End If

    ' Before you start filling the PivotTable sheets with data, clear the existing contents.
    Application.EnableEvents = False ' disable events
    destSheet.Cells.ClearContents
    destSheet2.Cells.ClearContents
    Application.EnableEvents = True ' enable events again

    destSheet.Cells(1, 1).Value = "WorkerName"
    destSheet.Cells(1, 2).Value = "DateShifts"
    destSheet.Cells(1, 3).Value = "NumberShifts"
    destSheet.Cells(1, 1).EntireRow.Font.Bold = True

    ' Setup headers for PivotTable2
    destSheet2.Cells(1, 1).Value = "WorkerName"
    destSheet2.Cells(1, 2).Value = "DateMonth"
    destSheet2.Cells(1, 3).Value = "DataHeader"
    destSheet2.Cells(1, 4).Value = "DataValue"
    destSheet2.Cells(1, 1).EntireRow.Font.Bold = True

    Dim destSheetRow As Long, destSheet2Row As Long
    destSheetRow = 2
    destSheet2Row = 2

    For i = 3 To lastRow
        Name = CStr(srcSheet.Cells(i, Chr(Asc(squadCol) + 2)).Value) ' Convert cell content to a string
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

Sub PivotDataUnique(ByVal wsNum As Long, ByVal lastRow As Long, ByVal squadCol As String)
    
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
    Set srcSheet = ThisWorkbook.Sheets(wsNum)
    
    ' Check if "WorkersStatus" sheet exists. If not, create it.
    Dim ws As Worksheet
    Dim sheetExists3 As Boolean
    sheetExists3 = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "WorkersStatus" Then
            sheetExists3 = True
        End If
    Next ws
    
    If Not sheetExists3 Then
        Set destSheet3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        destSheet3.Name = "WorkersStatus"
    Else
        Set destSheet3 = ThisWorkbook.Sheets("WorkersStatus")
    End If

    ' Clear PivotTable3 sheet before adding new data
    destSheet3.Cells.ClearContents

    GroupRelevant = MsgBox("Is the column before Squad relevant?", vbYesNo)
    
    If GroupRelevant = vbYes Then
        destSheet3.Cells(1, 1).Value = "WorkerGroup"
        destSheet3.Cells(1, 2).Value = "WorkerSquad"
        destSheet3.Cells(1, 3).Value = "SquadSymbol"
        destSheet3.Cells(1, 4).Value = "WorkerName"
    Else
        destSheet3.Cells(1, 1).Value = "WorkerSquad"
        destSheet3.Cells(1, 2).Value = "SquadSymbol"
        destSheet3.Cells(1, 3).Value = "WorkerName"
    End If

    destSheet3.Cells(1, 1).EntireRow.Font.Bold = True
    
    Dim destSheet3Row As Long
    destSheet3Row = 2

    Set dictUnique = CreateObject("Scripting.Dictionary")

    For i = 3 To lastRow
        Group = CStr(srcSheet.Cells(i, Chr(Asc(squadCol) - 1)).Value) ' Convert cell content to a string
        Squad = CStr(srcSheet.Cells(i, squadCol).Value) ' Convert cell content to a string
        Abbreviation = CStr(srcSheet.Cells(i, Chr(Asc(squadCol) + 1)).Value) ' Convert cell content to a string
        Name = CStr(srcSheet.Cells(i, Chr(Asc(squadCol) + 2)).Value) ' Convert cell content to a string


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

Sub CheckForEmptyCells()

    ' Check sheets "WorkersShifts", "WorkersMonthData", and "WorkersStatus" for empty cells
    Dim wsNames As Variant
    Dim ws As Worksheet
    Dim i As Long, j As Long, lastCol As Long
    Dim rowEmpty As Boolean
    Dim msg As String
    
    wsNames = Array("WorkersShifts", "WorkersMonthData", "WorkersStatus")
    
    For i = LBound(wsNames) To UBound(wsNames)
        Set ws = ThisWorkbook.Sheets(wsNames(i))
        ' Find the last header column
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        msg = ""
        For j = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            rowEmpty = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(j, 1), ws.Cells(j, lastCol))) = 0
            If rowEmpty Then
                Exit For
            Else
                If Application.WorksheetFunction.CountBlank(ws.Range(ws.Cells(j, 1), ws.Cells(j, lastCol))) > 0 Then
                    msg = msg & vbNewLine & "Row " & j & " in sheet " & wsNames(i)
                End If
            End If
        Next j
        If msg <> "" Then
            MsgBox "The following rows in " & wsNames(i) & " have empty cells:" & msg
        End If
        Set ws = Nothing
    Next i

End Sub




