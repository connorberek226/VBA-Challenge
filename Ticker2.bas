Attribute VB_Name = "Module2"
Sub Ticker2():
    
For Each ws In Worksheets
    
    ' Calculates Greatest, Worst and Greatest Total Volume
    ' of each stock.
    
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    
    ' For loop to run through the new data from Ticker() module
    ' to establish the Greatest Increase, Decrease and Volume
    ' out of all the stocks.
    For i = 2 To 3500
        
        If ws.Cells(i, 11).Value > GreatestIncrease Then
            ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
            GreatestIncrease = ws.Cells(i, 11).Value
            
            ' Populates cell with greatest increase percentage
            ws.Cells(2, 16).Value = Format(GreatestIncrease, "Percent")
        End If
        
        If ws.Cells(i, 11).Value < GreatestDecrease Then
            ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
            GreatestDecrease = ws.Cells(i, 11).Value
            
            ' Populates cell with greatest decrease percentage
            ws.Cells(3, 16).Value = Format(GreatestDecrease, "Percent")
            
        End If
        
        If ws.Cells(i, 12).Value > GreatestVolume Then
            ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
            GreatestVolume = ws.Cells(i, 12).Value
            
            ' Populates cell with greatest volume
            ws.Cells(4, 16).Value = GreatestVolume
        End If
    
    Next i
    
Next ws

End Sub
