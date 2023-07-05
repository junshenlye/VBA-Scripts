Sub GenerateCleanedFile()
    ' Set source worksheet
    Dim sourceWorksheet As Worksheet
    Set sourceWorksheet = ThisWorkbook.Worksheets("Jun23") ' Update with your source worksheet name
    
    ' Set mapping file path and open mapping workbook
    Dim mappingWorkbook As Workbook
    Dim mappingFilePath As String
    mappingFilePath = "E:\Thales\Sales report to Thales\Customer.xlsx" ' Update with your mapping file path
    Set mappingWorkbook = Workbooks.Open(mappingFilePath)
    
    ' Set mapping dictionary
    Dim mappingDict As Object
    Set mappingDict = CreateObject("Scripting.Dictionary")
    
    ' Load mapping data into dictionary
    Dim mappingData As Range
    Set mappingData = mappingWorkbook.Worksheets("Sheet1").Range("A1:B5")
    
    For Each mappingRow In mappingData.Rows
        Dim site As String
        Dim shortForm As String
        site = mappingRow.Cells(2).Value
        shortForm = mappingRow.Cells(1).Value
        mappingDict(site) = shortForm
    Next mappingRow

    ' Create new workbook and set reference to first worksheet
    Dim newWorkbook As Workbook
    Set newWorkbook = Workbooks.Add
    Dim newWorksheet As Worksheet
    Set newWorksheet = newWorkbook.Worksheets(1)
    
    ' Copy data from source worksheet to new worksheet
    sourceWorksheet.UsedRange.Copy newWorksheet.Cells(1, 1)
    ' Clear cell format starting from J2
    Dim InvoiceRange As Range
    Set InvoiceRange = newWorksheet.Range("J2").Resize(newWorksheet.Rows.Count - 1, newWorksheet.Columns.Count - 9) ' Adjust the "- 9" value if needed

    InvoiceRange.ClearFormats

    ' Get last row in new worksheet
    Dim lastRow As Long
    lastRow = newWorksheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Replace Thales DIS requestor site with short form version
    Dim thalesSiteCol As Range
    Set thalesSiteCol = newWorksheet.Rows(1).Find("Thales DIS requestor site")
    Dim thalesSiteHeader As String ' Store the header title
    thalesSiteHeader = thalesSiteCol.Value ' Get the header title
    thalesSiteCol = thalesSiteCol.Column ' Get the column number

    Dim thalesSiteCell As Range
    For Each thalesSiteCell In newWorksheet.Range(newWorksheet.Cells(2, thalesSiteCol), newWorksheet.Cells(lastRow, thalesSiteCol))
        If mappingDict.Exists(thalesSiteCell.Value) Then
            thalesSiteCell.Value = mappingDict(thalesSiteCell.Value)
        End If
    Next thalesSiteCell

    ' Set the header title back to the modified column
    newWorksheet.Cells(1, thalesSiteCol).Value = thalesSiteHeader

    ' Filter rows based on Sunningdale Plant and Thales DIS requestor site
    newWorksheet.AutoFilterMode = False ' Clear any existing filters
    newWorksheet.Range("A1").AutoFilter Field:=1, Criteria1:="<>"
    newWorksheet.Range("A1").AutoFilter Field:=thalesSiteCol.Column, Criteria1:="<>"

    ' Copy filtered data to a new worksheet "Summarized Version"
    Dim summarizedWorksheet As Worksheet
    Set summarizedWorksheet = newWorkbook.Sheets.Add(After:=newWorkbook.Sheets(newWorkbook.Sheets.Count))
    summarizedWorksheet.Name = "Summarized Version"

    newWorksheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Copy summarizedWorksheet.Range("A1")

    ' Format the data as a table in the "Summarized Version" worksheet
    Dim summarizedDataRange As Range
    Set summarizedDataRange = summarizedWorksheet.UsedRange
    Dim summarizedTable As ListObject
    Set summarizedTable = summarizedWorksheet.ListObjects.Add(xlSrcRange, summarizedDataRange, , xlYes)
    summarizedTable.TableStyle = "TableStyleMedium2"

    ' Find the table range for the summarized table
    Dim summarizedTableRange As Range
    Set summarizedTableRange = summarizedTable.Range
    
    ' Find the column numbers for the columns to be deleted
    Dim deleteColumns As Variant
    deleteColumns = Array("Currency", "Shipment Date (Invoice date)", "Form factor type", "PO number", "Unit price  Price / 1000 ")
    
    Dim deleteColumn As Variant
    For Each deleteColumn In deleteColumns
        Dim deleteColumnIndex As Variant
        deleteColumnIndex = Application.Match(deleteColumn, summarizedTableRange.Rows(1), 0)
    
        If Not IsError(deleteColumnIndex) Then
            summarizedWorksheet.Columns(deleteColumnIndex).Delete
        End If
    Next deleteColumn

    ' Sort and group the data in the "Summarized Version" worksheet
    Dim lastRowSummarized As Long
    lastRowSummarized = summarizedWorksheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Sort the data based on columns A, B, and C in ascending order
    With summarizedWorksheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=summarizedWorksheet.Range("A2:A" & lastRowSummarized), SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=summarizedWorksheet.Range("B2:B" & lastRowSummarized), SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=summarizedWorksheet.Range("C2:C" & lastRowSummarized), SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange summarizedWorksheet.Range("A1:E" & lastRowSummarized)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Group and sum the data
    Dim i As Long
    For i = lastRowSummarized To 3 Step -1
        If summarizedWorksheet.Cells(i, 1).Value = summarizedWorksheet.Cells(i - 1, 1).Value And summarizedWorksheet.Cells(i, 2).Value = summarizedWorksheet.Cells(i - 1, 2).Value And summarizedWorksheet.Cells(i, 3).Value = summarizedWorksheet.Cells(i - 1, 3).Value Then
            summarizedWorksheet.Cells(i - 1, 4).Value = summarizedWorksheet.Cells(i - 1, 4).Value + summarizedWorksheet.Cells(i, 4).Value
            summarizedWorksheet.Cells(i - 1, 5).Value = summarizedWorksheet.Cells(i - 1, 5).Value + summarizedWorksheet.Cells(i, 5).Value
            summarizedWorksheet.Rows(i).Delete
        End If
    Next i

    ' Save the new workbook
    Dim savePath As String
    savePath = "C:\Users\Jun Wu\Downloads\clean_ref.xlsx" ' Update with the desired output file path and name
    newWorkbook.SaveAs savePath
    
    ' Close the mapping workbook without saving changes
    mappingWorkbook.Close True
    
    MsgBox "Data cleaning and summarization complete. The cleaned file has been saved."
End Sub

