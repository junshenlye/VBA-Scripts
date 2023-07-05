Sub CrossReferenceData()
    ' Define the worksheets and file paths
    Dim compoundWorkbook As Workbook
    Dim salesWorkbook As Workbook
    Dim compoundWorksheet As Worksheet
    Dim salesWorksheet As Worksheet
    Dim compoundFilePath As String
    Dim salesFilePath As String
    
    ' Set the file paths
    compoundFilePath = "C:\Users\Jun Wu\Downloads\OneDrive_2023-07-03\Amortization report (Updated)\4. 2023 Q2 Mold Amortization Report-May-draft.xlsx"
    salesFilePath = "C:\Users\Jun Wu\Downloads\OneDrive_2023-07-03\site sales report Jun23\STCZ SAP Jun23 sales report-230703.XLSX"
    
    ' Open the compound report file
    Set compoundWorkbook = Workbooks.Open(compoundFilePath)
    Set compoundWorksheet = compoundWorkbook.Worksheets("Original - Internal") ' Update the worksheet name if needed
    
    ' Open the sales report file
    Set salesWorkbook = Workbooks.Open(salesFilePath)
    Set salesWorksheet = salesWorkbook.Worksheets("Output") ' Update the worksheet name if needed
    
    ' Find the last rows in the compound and sales reports
    compoundLastRow = compoundWorksheet.Cells(compoundWorksheet.Rows.Count, "D").End(xlUp).Row
    salesLastRow = salesWorksheet.Cells(salesWorksheet.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through the compound report
    For i = 5 To compoundLastRow
        ' Retrieve the Mold # value in the compound report
        compoundMold = compoundWorksheet.Cells(i, "D").Value
        
        ' Loop through the sales report
        For j = 1 To salesLastRow
            ' Retrieve the MoldSerial value in the sales report
            salesMoldSerial = salesWorksheet.Cells(j, "A").Value
            
            ' Check if the Mold # and MoldSerial match
            If compoundMold = salesMoldSerial Then
                ' Retrieve the Qty value from the sales report
                salesQty = salesWorksheet.Cells(j, "B").Value
                
                ' Update the corresponding cell in the compound report
                compoundWorksheet.Cells(i, "AR").Value = salesQty ' Assuming G is the column for SCTZ
                
                ' Exit the inner loop as the match has been found
                Exit For
            End If
        Next j
    Next i
    
    ' Save and close the files
    compoundWorkbook.Save
End Sub



