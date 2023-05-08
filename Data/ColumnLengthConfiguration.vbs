Sub ColumnsLengthConfigurationForSheet(sheetName As String)
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets(sheetName)

ws.Range("A1:B1").Interior.Color = RGB(54, 66, 51)
ws.Range("A14").Interior.Color = RGB(153, 0, 51)
ws.Range("C1").Interior.Color = RGB(90, 100, 125)
ws.Cells.Font.Bold = True
ws.Cells.Font.Size = 14
ws.Range("A2:A20").Font.Color = RGB(113, 110, 203)
ws.Columns("A").ColumnWidth = 25
ws.Range("A14").RowHeight = 50
ws.Columns("B").ColumnWidth = 110
ws.Columns("C").ColumnWidth = 75
ActiveWindow.Zoom = 75

End Sub