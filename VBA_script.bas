Attribute VB_Name = "Module1"
Sub ticker()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        'MsgBox (ws)
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        x = 2
        i = 2
        
            Do While ws.Cells(x, 1) <> ""
                tick = ws.Cells(x, 1)
                
               'Does some things on the first insance of a particular ticker
                If ws.Cells(x - 1, 1) <> ws.Cells(x, 1) Then
                    'Define the first 'open' value
                    first = ws.Cells(x, 3)
                    End If
                    
                'Does some things on the last instance
                If ws.Cells(x + 1, 1) <> ws.Cells(x, 1) Then
                    
                    'Calculate some values
                    ycount = first - ws.Cells(x, 6).Value
                    pcount = ycount / first
                    
                    'Put some values/colour in some cells
                    ws.Cells(i, 9).Value = tick
                    
                    ws.Cells(i, 10).Value = ycount
                    If ycount > 0 Then
                        ws.Cells(i, 10).Interior.ColorIndex = 4
                    ElseIf ycount <= 0 Then
                        ws.Cells(i, 10).Interior.ColorIndex = 3
                    End If
                    
                    ws.Cells(i, 11).Value = FormatPercent(pcount, 2)
                    ws.Cells(i, 12).Value = vcount
                    
                    'Reset some values and go to the next row of the results table
                    vcount = 0
                    i = i + 1
                End If
                
        
                
                'Percent Change
                
                
                vcount = vcount + ws.Cells(x, 7)
                
                x = x + 1
            Loop
            
        'Calculate some stats
        increase = 0
        decrease = 0
        vol = 0
        i = 2
        Do While ws.Cells(i, 9) <> ""
            If ws.Cells(i, 11) > increase Then
                increase = ws.Cells(i, 11).Value
                ti = Cells(i, 9).Value
            End If
            If ws.Cells(i, 11) < decrease Then
                decrease = ws.Cells(i, 11).Value
                td = ws.Cells(i, 9)
            End If
            If ws.Cells(i, 12) > vol Then
                vol = ws.Cells(i, 12).Value
                tvol = ws.Cells(i, 9)
            End If
            i = i + 1
        Loop
        'Display the stats
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(2, 15).Value = ti
        ws.Cells(3, 15).Value = td
        ws.Cells(4, 15).Value = tvol
        
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 16).Value = FormatPercent(increase, 2)
        ws.Cells(3, 16).Value = FormatPercent(decrease, 2)
        ws.Cells(4, 16).Value = vol
        
        Next ws
            
End Sub


