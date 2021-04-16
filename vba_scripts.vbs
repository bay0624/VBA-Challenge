Sub Stocks()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Select
        Call StockMarket
    Next
    Application.ScreenUpdating = True
End Sub

Sub StockMarket()

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Range("O2").Value = "Greatest % increase"
Range("O3").Value = "Greatest % decrease"
Range("O4").Value = "Greatest total volume"

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"


Dim lastrow As Double

Dim ticker_name As String

Dim total_stock As Double

Dim table_row As Integer
table_row = 2

Dim open_price As Double

Dim close_price As Double

Dim yearly_change As Double

Dim percent_change As Double

Dim count As Integer
Dim new_count As Integer

lastrow = Cells(Rows.count, 1).End(xlUp).Row

For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker_name = Cells(i, 1).Value
        Range("I" & table_row).Value = ticker_name
        
        total_stock = total_stock + Cells(i, 7).Value
        Range("L" & table_row).Value = total_stock
        
        count = Application.WorksheetFunction.CountIf(Range("$A$2:$A$" & lastrow), Cells(i, 1).Value)
        new_count = count - 1
        
        close_price = Cells(i, 6).Value
        open_price = Cells(i - new_count, 3).Value
        yearly_change = close_price - open_price
        Range("J" & table_row).Value = yearly_change
        
        If open_price = 0 Then
            percent_change = 0
            Range("K" & table_row).Value = Format(percent_change, "Percent")
        Else
            percent_change = yearly_change / open_price
            Range("K" & table_row).Value = Format(percent_change, "Percent")
        End If
        
        table_row = table_row + 1
        
        total_stock = 0
    
    Else
        total_stock = total_stock + Cells(i, 7).Value

    End If

Next i

lastrow_2 = Cells(Rows.count, 10).End(xlUp).Row

For j = 2 To lastrow_2
    
    If Cells(j, 10).Value > 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
        
    Else
        Cells(j, 10).Interior.ColorIndex = 3
    
    End If
    
Next j

lastrow_3 = Cells(Rows.count, 10).End(xlUp).Row

max_percent = Application.WorksheetFunction.Max(Range("K:K"))
Cells(2, 17).Value = Format(max_percent, "Percent")

For x = 2 To lastrow_3
    If Cells(x, 11).Value = Cells(2, 17).Value Then
        Cells(2, 16).Value = Cells(x, 9).Value
    End If

Next x
    
min_percent = Application.WorksheetFunction.Min(Range("K:K"))
Cells(3, 17).Value = Format(min_percent, "Percent")

For y = 2 To lastrow_3
    If Cells(y, 11).Value = Cells(3, 17).Value Then
        Cells(3, 16).Value = Cells(y, 9).Value
    End If

Next y

total_vol = Application.WorksheetFunction.Max(Range("L:L"))
Cells(4, 17).Value = total_vol

For v = 2 To lastrow_3
    If Cells(v, 12).Value = Cells(4, 17).Value Then
        Cells(4, 16).Value = Cells(v, 9).Value
    End If

Next v

End Sub
