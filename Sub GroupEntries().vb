Sub GroupEntries()
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim lastRow As Long
    Dim moldSerialRange As Range
    Dim moldSerialCell As Range
    Dim currentSerial As String
    Dim entryRange As Range
    Dim nextEmptyRow As Long
    
    ' Set the source sheet and target sheet
    Set sourceSheet = ThisWorkbook.Sheets("Sheet1") ' Replace "Sheet1" with the name of your source sheet
    Set targetSheet = ThisWorkbook.Sheets.Add(After:=sourceSheet)
    targetSheet.Name = "entries"
    
    ' Set the mold serial range
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "BH").End(xlUp).Row
    Set moldSerialRange = sourceSheet.Range("BH2:BH" & lastRow)
    
    ' Sort the mold serial range
    moldSerialRange.Sort key1:=moldSerialRange, order1:=xlAscending, Header:=xlNo
    
    ' Initialize the next empty row in the target sheet
    nextEmptyRow = 1
    
    ' Loop through the mold serial numbers
    For Each moldSerialCell In moldSerialRange
        If moldSerialCell.Value <> "" Then
            ' Check if the current mold serial number is different from the previous one
            If moldSerialCell.Value <> currentSerial Then
                ' Add empty space before the new group
                If nextEmptyRow > 1 Then
                    targetSheet.Cells(nextEmptyRow, 1).EntireRow.Insert xlShiftDown
                    nextEmptyRow = nextEmptyRow + 1
                End If
                
                ' Update the current mold serial number
                currentSerial = moldSerialCell.Value
            End If
            
            ' Set the entry range as the current row
            Set entryRange = sourceSheet.Range("A" & moldSerialCell.Row & ":BH" & moldSerialCell.Row)
            
            ' Copy the entry range to the target sheet
            entryRange.Copy targetSheet.Cells(nextEmptyRow, 1)
            nextEmptyRow = nextEmptyRow + 1
        End If
    Next moldSerialCell
End Sub
