Attribute VB_Name = "Module1"
Sub StockDataCalculation()

'loop through sheets

Dim WS As Worksheet

For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    last_row = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
'make summary chart headings row'

Cells(1, "I").Value = "Ticker"
Cells(1, "J").Value = "Yearly Change"
Cells(1, "K").Value = "Percent Change"
Cells(1, "L").Value = "Total Stock Volume"

'make dim statements for variables to be analyzed

Dim ticker As String
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim volume As Double
    volume = 0
Dim summary_table_row As Single
    summary_table_row = 2
Dim i As Long

'calculate initial opening price

opening_price = Cells(2, 3).Value

'create loop to determine ticker changes

For i = 2 To last_row

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
            Range("I" & summary_table_row).Value = ticker
        closing_price = Cells(i, 6).Value
        yearly_change = closing_price - opening_price
            Range("J" & summary_table_row).Value = yearly_change

'calculate percent change
'insert conditional to avoid division by 0 error

        If (opening_price = 0 And closing_price = 0) Then
            percent_change = 0
        ElseIf (opening_price = 0 And closing_price <> 0) Then
            percent_change = 1
        Else
            percent_change = yearly_change / opening_price
                Range("K" & summary_table_row).Value = percent_change
                Range("K" & summary_table_row).NumberFormat = "0.00%"
        End If
        
'add volume

        volume = volume + Cells(i, 7).Value
            Range("L" & summary_table_row).Value = volume
            
'complete summary table

        summary_table_row = summary_table_row + 1
        
'reset opening price for all beyond initial ticker

        opening_price = Cells(i + 1, 3)
        
'reset volume

        volume = 0
    Else
        volume = volume + Cells(i, 7).Value
    End If
Next i

'now find last row of yearly_change for each WS

yearly_change_finalrow = WS.Cells(Rows.Count, 10).End(xlUp).Row

'color in cells by loss or gain via conditional loop

For j = 2 To yearly_change_finalrow
    If Cells(j, 10).Value >= 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
    ElseIf Cells(j, 10) < 0 Then
        Cells(j, 10).Interior.ColorIndex = 3
    End If
Next j

Next WS

End Sub

