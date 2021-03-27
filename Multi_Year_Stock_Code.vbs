Sub Stock():

Dim ticker As String
Dim totVolume As LongLong
totVolume = 0

Dim AnnualClose As Double

For Each ws In Worksheets
    
    Dim Summary_Table As Integer
        Summary_Table = 2
    
    Dim AnnualOpen As Double
        AnnualOpen = ws.Cells(2, 3).Value
   
    Dim YearlyChange As Double
    
    Dim PercentChange As Double
    

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Activate
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Value"
    
    For i = 2 To lastrow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            ticker = ws.Cells(i, 1).Value
        
            totVolume = totVolume + ws.Cells(i, 7).Value
            
            AnnualClose = ws.Cells(i, 6).Value
            
            YearlyChange = AnnualClose - AnnualOpen
                If AnnualOpen <> 0 Then
            
                    PercentChange = YearlyChange / AnnualOpen
            
                End If
            
            ws.Range("J" & Summary_Table).Value = YearlyChange
            
                If ws.Range("J" & Summary_Table).Value > 0 Then
                    ws.Range("J" & Summary_Table).Interior.ColorIndex = 4
                    
                    ElseIf ws.Range("J" & Summary_Table).Value < 0 Then
                        ws.Range("J" & Summary_Table).Interior.ColorIndex = 3
                            
                End If
                
            ws.Range("L" & Summary_Table).Value = totVolume
    
            ws.Range("I" & Summary_Table).Value = ticker
            
            ws.Range("K" & Summary_Table).Value = PercentChange
            ws.Range("K" & Summary_Table).NumberFormat = "0.00%"
        
            Summary_Table = Summary_Table + 1
        
            totVolume = 0
            
            YearlyChange = 0
            
            AnnualOpen = ws.Cells(i + 1, 3).Value
                
            Else
        
            totVolume = totVolume + ws.Cells(i, 7).Value
                
        End If
    
    Next i
    ws.Cells(1, 16).Value = "Ticker"
    
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastrow)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & lastrow)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & lastrow))
    
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), 0)
    
    Range("P2") = Cells(increase_number + 1, 9)
    Range("P3") = Cells(decrease_number + 1, 9)
    Range("P4") = Cells(volume_number + 1, 9)
    
    ws.Cells.Columns.AutoFit
    
Next ws



End Sub


