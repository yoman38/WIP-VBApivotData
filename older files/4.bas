Attribute VB_Name = "Module2"
Sub PivotData()

    Dim srcSheet As Worksheet
    Dim destSheet As Worksheet
    Dim i As Long, j As Long
    Dim lastRow As Long
    Dim name As String
    Dim monthYear As String
    Dim day As Variant
    Dim shift As String
    Dim month As Integer
    Dim year As String
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

    ' Check if "PivotTable" sheet exists. If not, create it.
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    sheetExists = False
    For Each ws In ThisWorkbook.Sheets
        If ws.name = "PivotTable" Then
            sheetExists = True
            Exit For
        End If
    Next ws

    If Not sheetExists Then
        Set destSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        destSheet.name = "PivotTable"
    Else
        Set destSheet = ThisWorkbook.Sheets("PivotTable")
    End If
    
    lastRow = InputBox("Please enter the last row number:", "Input needed", "1000")

    destSheet.Cells(1, 1).Value = "Name"
    destSheet.Cells(1, 2).Value = "Date"
    destSheet.Cells(1, 3).Value = "Shift"
    destSheet.Cells(1, 1).EntireRow.Font.Bold = True
    
    Dim destSheetRow As Long
    destSheetRow = 2

    For i = 3 To lastRow Step 2
        name = srcSheet.Cells(i, "G").Value
        monthYear = srcSheet.Cells(i, "H").Value
        
        ' Skip if name is empty, equals 'Nazwisko i imiê', or equals '-'
        If name = "" Or name = "Nazwisko i imiê" Or name = "-" Then
            GoTo NextRow
        End If
        
        ' Split monthYear string into month and year
        If InStr(monthYear, "zm.") > 0 Then
            month = monthDict(Split(monthYear, " zm. ")(0))
            year = Split(monthYear, " zm. ")(1)
        Else
            month = monthDict(Split(monthYear, " ")(0))
            year = Split(monthYear, " ")(1)
        End If
        
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
        
NextRow:
    Next i
    
    destSheet.Columns("B:B").NumberFormat = "yyyy-mm-dd"
    
End Sub

