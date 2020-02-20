Attribute VB_Name = "Module1"
Sub homework()


Dim ws As Worksheet
Dim ticker As String
Dim volume As Long
Dim open_year As Double
Dim close_year As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer

On Error Resume Next

For Each ws In ThisWorkbook.Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
    Summary_Table_Row = 2

    volume = 0

        For i = 2 To ws.UsedRange.Rows.Count
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
         
            ticker = ws.Cells(i, 1).Value
            volume = ws.Cells(i, 7).Value
            Sum = volume + Sum

            open_year = ws.Cells(i, 3).Value
            close_year = ws.Cells(i, 6).Value

            yearly_change = close_year - open_year
            percent_change = (close_year - open_year) / close_year

            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            ws.Cells(Summary_Table_Row, 11).Value = percent_change
            ws.Cells(Summary_Table_Row, 12).Value = Sum
            Summary_Table_Row = Summary_Table_Row + 1
            Sum = 0

            Else
            volume = ws.Cells(i, 7).Value
            Sum = volume + Sum
        
        End If


    Next i
    
ws.Columns("K").NumberFormat = "0.00%"


    Dim rg As Range
    Dim g As Long
    Dim c As Long
    Dim color_cell As Range
    
    Set rg = ws.Range("J2", Range("J2").End(xlDown))
    c = rg.Cells.Count
    
    For g = 1 To c
    Set color_cell = rg(g)
    Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.Color = vbGreen
            End With
        Case Is < 0
            With color_cell
                .Interior.Color = vbRed
            End With
       End Select
    Next g

Next ws

End Sub



