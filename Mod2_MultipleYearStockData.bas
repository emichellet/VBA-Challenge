Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock_Data():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        'Current row
        Dim i As Long
        'Start row of ticker block
        Dim j As Long
        'Index counter to fill Ticker row
        Dim TickCount As Long
        'Last row column A
        Dim LastRowA As Long
        'Last row columnI
        Dim LastRowI As Long
        'Variable for percent change calculation
        Dim PercentChange As Double
        'Variable for greatest increase calculation
        Dim GreatIncrease As Double
        'Variable for greatest decrease calculation
        Dim GreatDecrease As Double
        'Variable for greatest total volume
        Dim GreatVolume As Double
        
        'Get WorksheeName
        WorksheetName = ws.Name
        
        'Create some column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        
        'Set the ticker counter to the first row
        TickCount = 2
        
        'Set start row to 2
        j = 2
        
        'Find the last non-blank cell in Column A
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'MsgBox("Last row in Column A is " & LastRowA)
        
            'Loop through all of the rows
            For i = 2 To LastRowA
            
                'Check if the ticker name has changed
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Write ticker in column I (Column #9)
                ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                
                'Calculate and write the Yearly Change in Column j (Column #10)
                ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                'Do some conditional formatting
                If ws.Cells(TickCount, 10).Value < 0 Then
                
                'Set the cell background color to a dark red
                ws.Cells(TickCount, 10).Interior.ColorIndex = 9
                
                Else
                
                'Set the cell background color to a light green
                ws.Cells(TickCount, 10).Interior.ColorIndex = 35
                
                End If
                
                'Calculate and write the percent change in column K (Column #11)
                If ws.Cells(j, 3).Value <> 0 Then
                PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                
                'Formatting for percent
                ws.Cells(TickCount, 11).Value = Format(PercentChange, "Percent")
                
                Else
                
                ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                
                End If
                
                'Calculate and write the total volume within Column L (Column #12)
                ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase the TickCount by 1
                TickCount = TickCount + 1
                
                'Set a new start row of the ticker block
                j = i + 1
                
                End If
                
            Next i
        
        'Find last non-blank cell in Column I (Column #9)
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox("Last row in column I is " & LastRowI)
        
        'Prepare for a summary
        GreatVolume = ws.Cells(2, 12).Value
        GreatIncrease = ws.Cells(2, 11).Value
        GreatDecrease = ws.Cells(2, 11).Value
        
            'Loop for the summary
            For i = 2 To LastRowI
            
                'For the greatest total volume, check if the next value is larger; if yes, take over a new value and populate ws.Cells
                    If ws.Cells(i, 12).Value > GreatVolume Then
                    GreatVolume = ws.Cells(i, 12).Value
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                    
                    Else
                    
                    GreatVolume = GreatVolume
                    
                    End If
                    
                'For the greatest increase, check if the next value is larger; if yes, take over a new value and populate ws.Cells
                    If ws.Cells(i, 11).Value > GreatIncrease Then
                    GreatIncrease = ws.Cells(i, 11).Value
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                    
                    Else
                    
                    GreatIncrease = GreatIncrease
                    
                    End If
                    
                'For the greatest decrease, check if the next value is smaller; if it is, take over a new value and populate the cells
                    If ws.Cells(i, 11).Value < GreatDecrease Then
                    GreatDecrease = ws.Cells(i, 11).Value
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                    
                    Else
                    
                    GreatDecrease = GreatDecrease
                    
                    End If
                    
            'Write the summary results in ws.Cells
            ws.Cells(2, 17).Value = Format(GreatIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVolume, "Scientific")
            
            Next i
            
        'Adjust the column width of the worksheets automatically
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
    
    Next ws
        
                           
End Sub
