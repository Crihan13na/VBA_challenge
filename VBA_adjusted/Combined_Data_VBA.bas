Attribute VB_Name = "Module1"
Sub stocks()
    
    ' Add a sheet named "Combined Data"
    Sheets.Add.Name = "Combined_Data"
    'move created sheet to be first sheet
    Sheets("Combined_Data").Move Before:=Sheets(1)
    ' Specify the location of the combined sheet
    Set combined_sheet = Worksheets("Combined_Data")

    ' Loop through all sheets
    For Each ws In Worksheets

        ' Find the last row of the combined sheet after each paste
        ' Add 1 to get first empty row
        LastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1

        ' Find the last row of each worksheet
        ' Subtract one to return the number of rows without header
        LastRowstock = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1

        ' Copy the contents of each state sheet into the combined sheet
        combined_sheet.Range("A" & LastRow & ":G" & ((LastRowstock - 1) + LastRow)).Value = ws.Range("A2:G" & (LastRowstock + 1)).Value
Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock As Variant
total_stock = 0
Dim close_price As Double
Dim open_price As Double
Dim summary_table_row As Integer
summary_table_row = 2
Dim j As Integer
Dim i As Long
Dim WorksheetsName As String

'set headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "total stock"

'set the price
    open_price = ws.Cells(2, 3).Value

     'Loop
For i = 2 To LastRowstock
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     ticker = ws.Cells(i, 1).Value
   
        'set close price
        close_price = ws.Cells(i, 6).Value
        yearly_change = close_price - open_price
        'Add %_change
    If (open_price = 0 And close_price = 0) Then
        percent_change = 0

ElseIf (open_price = 0 And close_price <> 0) Then
        percent_change = 1
    Else
        percent_change = yearly_change / open_price
        ws.Cells(i, 11).NumberFormat = "0.00%"
    End If
        'Add total volume of stock
        total_stock = total_stock + Cells(i, 7).Value
   
    ws.Range("I" & summary_table_row).Value = ticker
    ws.Range("J" & summary_table_row).Value = yearly_change
    ws.Range("K" & summary_table_row).Value = percent_change
    ws.Range("L" & summary_table_row).Value = total_stock
   
        summary_table_row = summary_table_row + 1
        i = i + 1
    'reset the price and the total volume
        open_price = ws.Cells(i + 1, 3)
        total_stock = 0
   Else
        total_stock = total_stock + ws.Cells(i, 7).Value
     End If
Next i

     
      'set the last row of each ws
      wsLastRowstock = ws.Cells(Rows.Count, 9).End(xlUp).Row
      'Add Color for positiv % change and negativ % change
     For j = 2 To wsLastRowstock
     If (ws.Cells(j, 10).Value > 0 Or ws.Cells(j, 10).Value = 0) Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
     ElseIf ws.Cells(j, 10).Value < 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
    End If
Next j

'set variables for the hard solution
ws.Cells(2, 15).Value = "greatest%increase"
ws.Cells(3, 15).Value = "greatest%decrease"
ws.Cells(4, 15).Value = "greatestTotalVol"
ws.Cells(1, 16).Value = "ticker"
ws.Cells(1, 17).Value = "total_stock"

'Loop
For i = 2 To wsLastRowstock
If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & wsLastRowstock)) Then
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
    ws.Cells(2, 17).NumberFormat = "0.00%"
ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & wsLastRowstock)) Then
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
    ws.Cells(3, 17).NumberFormat = "0.00%"
ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & wsLastRowstock)) Then
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
    End If
Next i


Next ws

    ' Copy the headers from sheet 1
    combined_sheet.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value
    
    ' Autofit to display data
    combined_sheet.Columns("A:G").AutoFit
End Sub


