Attribute VB_Name = "Module1"
Sub Stocks()

Dim ticker As String
Dim tickers As Integer
Dim yearly As Double
Dim opening As Double
Dim closing As Double
Dim percent As Double
Dim stock_volume As Double
Dim lastrow As Long

For Each ws In Worksheets

    ws.Activate
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    tickers = 0
    ticker = ""
    yearly = 0
    opening = 0
    percent = 0
    stock_volume = 0
    
For i = 2 To lastrow
    
    ticker = Cells(i, 1).Value
    
    If opening = 0 Then
            opening = Cells(i, 3).Value
        End If
    
    stock_volume = stock_volume + Cells(i, 7).Value
    
    If Cells(i + 1, 1).Value <> ticker Then
    
    tickers = tickers + 1
    
    Cells(tickers + 1, 9) = ticker
    
    closing = Cells(i, 6)
    
    yearly = closing - opening
    
    Cells(tickers + 1, 10).Value = yearly
        
        
        If yearly < 0 Then
                Cells(tickers + 1, 10).Interior.ColorIndex = 3
        
        
        ElseIf yearly > 0 Then
                Cells(tickers + 1, 10).Interior.ColorIndex = 4
        
        End If
            
        If opening = 0 Then
                percent = 0
            Else
                percent = (yearly / opening)
            End If
            
            Cells(tickers + 1, 11).Value = Format(percent, "Percent")
            
            opening = 0
            
            Cells(tickers + 1, 12).Value = stock_volume
            
            stock_volume = 0
        
        End If
        
    Next i
        
 Next ws
   

       
    
End Sub

