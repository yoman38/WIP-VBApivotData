Attribute VB_Name = "Module2"
Sub PivotData()

    Dim srcSheet As Worksheet
    Dim destSheet As Worksheet, destSheet2 As Worksheet
    Dim i As Long, j As Long, k As Long
    Dim lastRow As Long
    Dim name As String
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
        If ws.name = "PivotTable" Then
            sheetExists = True
        ElseIf ws.name = "PivotTable2" Then
            sheetExists2 = True
        End If
    Next ws

    If Not sheetExists Then
        Set destSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        destSheet.name = "PivotTable"
    Else
        Set destSheet = ThisWorkbook.Sheets("PivotTable")
    End If

    If Not sheetExists2 Then
        Set destSheet2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        destSheet2.name = "PivotTable2"
    Else
        Set destSheet2 = ThisWorkbook.Sheets("PivotTable2")
    End If

    lastRow = InputBox("Please enter the last row number:", "Input needed", "1000")

    destSheet.Cells(1, 1).Value = "Name"
    destSheet.Cells(1, 2).Value = "Date"
    destSheet.Cells(1, 3).Value = "Shift"
    destSheet.Cells(1, 1).EntireRow.Font.Bold = True

    ' Copy headers from source sheet to PivotTable2 for columns AT to BF
    For k = 46 To 58
        destSheet2.Cells(1, k - 44).Value = srcSheet.Cells(2, k).Value
    Next k
    destSheet2.Cells(1, 1).EntireRow.Font.Bold = True

    Dim destSheetRow As Long, destSheet2Row As Long
    destSheetRow = 2
    destSheet2Row = 2

    For i = 3 To lastRow
        name = CStr(srcSheet.Cells(i, "G").Value) ' Convert cell content to a string
        monthYear = srcSheet.Cells(i, "H").Value
        
        ' Skip if name is empty, equals 'Nazwisko i imiê', equals '-', or equals '0'
        If name = "" Or name = "Nazwisko i imiê" Or name = "-" Or name = "0" Then
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
                destSheet2.Cells(destSheet2Row, 1).Value = DateSerial(year, month, 1)
                destSheet2.Cells(destSheet2Row, k - 44).Value = data
            End If
        Next k
        destSheet2Row = destSheet2Row + 1

        ' Add data to PivotTable for columns I to AS
        For j = 9 To 44
            day = srcSheet.Cells(i, j).Value
            shift = srcSheet.Cells(i + 1, j).Value
            
            ' Only process cells that contain a numeric day value
            If IsNumeric(day) Then
                
                If Not IsEmpty(shift) And Not IsError(shift) Then
                    destSheet.Cells(destSheetRow, 1).Value = name
                    destSheet.Cells(destSheetRow, 2).Value = DateSerial(year, month, day)
                    destSheet.Cells(destSheetRow, 3).Value = shift
                    destSheetRow = destSheetRow + 1
                End If
                
            End If
            
        Next j
        i = i + 1 ' Skip next row as it contains shifts for current row's dates

NextRow:
    Next i

    destSheet.Columns("B:B").NumberFormat = "yyyy-mm"
    destSheet2.Columns("A:A").NumberFormat = "yyyy-mm"

End Sub

