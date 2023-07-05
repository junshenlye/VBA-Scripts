Sub GenerateMoldSerialAndSortByQty()
    Dim lastRow As Long
    Dim i As Long
    
    ' Set the worksheet and column variables
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Replace "Sheet1" with the actual sheet name
    Dim batchColumn As Range
    Set batchColumn = ws.Range("AZ2:AZ" & ws.Cells(ws.Rows.Count, "AZ").End(xlUp).Row)
    
    ' Add a new column for MoldSerial
    ws.Columns("BH:BH").Insert Shift:=xlToRight
    ws.Cells(1, "BH").Value = "MoldSerial"
    
    ' Loop through each row in the batch column
    lastRow = ws.Cells(ws.Rows.Count, "AZ").End(xlUp).Row
    For i = 2 To lastRow
        Dim batchValue As String
        batchValue = batchColumn.Cells(i - 1).Value
        
        ' Extract the last few digits until it reaches an alphabetic character
        Dim lastDigits As String
        Dim j As Long
        For j = Len(batchValue) To 1 Step -1
            If Not IsNumeric(Mid(batchValue, j, 1)) Then
                Exit For
            End If
        Next j
        lastDigits = Mid(batchValue, j + 1)
        
        ' Create the MoldSerial value
        Dim moldSerial As String
        moldSerial = "SMS" & lastDigits
        
        ' Write the MoldSerial value in the new column
        ws.Cells(i, "BH").Value = moldSerial
    Next i
    
    ' Set the range of your columns
    Dim serialNumColumn As Range, qtyColumn As Range
    Dim qtyDict As Object, key As Variant
    Dim totalQty As Long
    Set qtyColumn = ws.Range("S2:S" & ws.Range("S" & ws.Rows.Count).End(xlUp).Row)
    Set serialNumColumn = ws.Range("BH2:BH" & ws.Range("BH" & ws.Rows.Count).End(xlUp).Row)
    
    ' Create a dictionary to store the total quantity for each unique serial number
    Set qtyDict = CreateObject("Scripting.Dictionary")
    
    ' Loop through the serial number column and sum up the quantities
    For Each cell In serialNumColumn
        If Not qtyDict.exists(cell.Value) Then
            qtyDict.Add cell.Value, 0
        End If
        qtyDict(cell.Value) = qtyDict(cell.Value) + qtyColumn(cell.Row - 1, 1).Value
    Next cell
    
    ' Clear the existing data in the output sheet
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Output"
    Sheets("Output").Range("A1:B" & Sheets("Output").Cells(Rows.Count, 1).End(xlUp).Row).ClearContents
    
    ' Sort the serial numbers based on the total quantity in descending order
    lastRow = 1
    For Each key In qtyDict
        totalQty = qtyDict(key)
        Sheets("Output").Range("A" & lastRow).Value = key
        Sheets("Output").Range("B" & lastRow).Value = totalQty
        lastRow = lastRow + 1
    Next key
    
    Sheets("Output").Range("A1:B" & lastRow - 1).Sort key1:=Sheets("Output").Range("B1"), order1:=xlDescending, Header:=xlYes
    
    ' Cleanup
    Set qtyDict = Nothing
    Set qtyColumn = Nothing
    Set serialNumColumn = Nothing
End Sub
