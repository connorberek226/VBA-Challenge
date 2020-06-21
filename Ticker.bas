Attribute VB_Name = "Module1"
Sub Ticker():

For Each ws In Worksheets


    ' Denoting variables to be called upon during the for loop.
    
    Dim lastrow As Long
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim TickerName As String
    
    Dim TotalVolume As Double
    
    Dim YearlyChange As Double
    
    Dim PercentChange As Double
    
    Dim OpeningPrice As Double
        
        
    TotalVolume = 0
    
    Dim TickerRow As Integer
    
    TickerRow = 2
    
    ' Setting the value of the opening price to hold for the first run through
    ' the loop and then it will be replaced with a new value every time.
    OpeningPrice = ws.Cells(2, 3).Value
    
    For i = 2 To lastrow
        
        ' Checking to see if the next cell in the row is the same
        ' current row. If not, it enters into the "If/Then" statement
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Ticker name is defined
                TickerName = ws.Cells(i, 1).Value
            
                ' Total volume is established
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
                ' Yearly change in stock price is set
                YearlyChange = ws.Cells(i, 6).Value - OpeningPrice
                
                ' Checks to see if opening price is non-zero
                ' enters "If/Then" statement if True.
                If OpeningPrice > 0 Then
                    PercentChange = YearlyChange / OpeningPrice
                    Else
                        PercentChange = 0
                End If
                
                ' Prints out data in necessary cells for each stock,
                ' volume total, percent change and yearly change
                ws.Range("I" & TickerRow).Value = TickerName
            
                ws.Range("L" & TickerRow).Value = TotalVolume
            
                ws.Range("J" & TickerRow).Value = YearlyChange
            
                ws.Range("K" & TickerRow).Value = Format(PercentChange, "Percent")
            
            
            ' Resets variables
            TickerRow = TickerRow + 1
            
            TotalVolume = 0
            
            YearlyChange = 0
            
            PercentChange = 0
            
            OpeningPrice = ws.Cells(i + 1, 3).Value
        
        Else
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
        End If
        
    
    Next i

Next ws


End Sub
