Attribute VB_Name = "Module1"
Sub FinalStockAnalysis():

    For Each ws In Worksheets
    
        Dim i As Long
        Dim j As Long
        Dim TickCount As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim PercentCh As Double
        Dim GreatInc As Double
        Dim GreatDec As Double
        Dim GreatVol As Double
        
        'Column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "% Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Set Ticker Counter to first row
        TickCount = 2
        
        'Set start row to 2
        j = 2
        
        'Find last row in column A
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Loop through all rows
            For i = 2 To LastRowA
            
                'Check for next ticker
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Write ticker in Col I
                ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                
                'Write yearly change in Col J
                ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'Conditional formatting for cell background. Red=neg, Green=pos
                    If ws.Cells(TickCount, 10).Value < 0 Then
                    
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                
                    Else
                
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculate and write % change in Col K
                    If ws.Cells(j, 3).Value <> 0 Then
                    PercentCh = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    ws.Cells(TickCount, 11).Value = Format(PercentCh, "Percent")
                    
                    Else
                    
                    ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'Calculate and write total volume in Col L
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase TickCount by 1
                TickCount = TickCount + 1
                
                'Set new start row for ticker
                j = i + 1
                
                End If
            
            Next i
            
        'Find last non-blank cell in column I
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Summary cells
        GreatVol = ws.Cells(2, 12).Value
        GreatInc = ws.Cells(2, 11).Value
        GreatDec = ws.Cells(2, 11).Value
        
            'Loop for summary
            For i = 2 To LastRowI
            
                'Greatest Total Volume
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVol = GreatVol
                
                End If
                
                'Greatest Increase
                If ws.Cells(i, 11).Value > GreatInc Then
                GreatInc = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatInc = GreatInc
                
                End If
                
                'Greatest decrease
                If ws.Cells(i, 11).Value < GreatDec Then
                GreatDec = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDec = GreatDec
                
                End If
                
            'Write summary results
            ws.Cells(2, 17).Value = Format(GreatInc, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDec, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
            Next i
            
        'Autofit columns
        Worksheets(ws.Name).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub


End Sub
