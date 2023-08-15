# module-1-challenge-vba
module 2 challenge

''' vba code

Sub stock_analysis():
   
    Dim total As Double
    Dim rowI As Long
    Dim change As Double
    Dim columnI As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Single
    Dim averageChange As Double
    Dim ws As Worksheet
    

    For Each ws In Worksheets
        columnI = 0
        total = 0
        change = 0
        start = 2
        dailyChange = 0
        
        

        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Change"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For rowI = 2 To rowCount
        
            If ws.Cells(rowI + 1, 1).Value <> ws.Cells(rowI, 1).Value Then
            
                total = total + ws.Cells(rowI, 7).Value
                
                If total = 0 Then
                
                    ws.Range("I" & 2 + columnI).Value = Cells(rowI, 1).Value
                    ws.Range("J" & 2 + columnI).Value = 0
                    ws.Range("K" & 2 + columnI).Value = "%" & 0
                    ws.Range("L" & 2 + columnI).Value = 0
                Else
                    If ws.Cells(start, 3) = 0 Then
                        For find_value = start To rowI
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If
                    
                    change = (ws.Cells(rowI, 6) - ws.Cells(start, 3))
                    percentChange = change / ws.Cells(start, 3)
                    
                    start = rowI + 1
                    
                    ws.Range("I" & 2 + columnI) = ws.Cells(rowI, 1).Value
                    ws.Range("J" & 2 + columnI) = change
                    ws.Range("J" & 2 + columnI).NumberFormat = "0.00"
                    ws.Range("K" & 2 + columnI).Value = percentChange
                    ws.Range("K" & 2 + columnI).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + columnI).Value = total
                    
                    
                    Select Case change
                        Case Is > 0
                            ws.Range("J" & 2 + columnI).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + columnI).Interior.ColorIndex = 3
                        Case Is > 0
                            ws.Range("J" & 2 + columnI).Interior.ColorIndex = 0
                    End Select
                         
                
                End If
            
                total = 0
                change = 0
                columnI = columnI + 1
                days = 0
                dailyChange = 0
                
            Else
                total = total + ws.Cells(rowI, 7).Value
                
            
            End If
            
        Next rowI
        
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
        
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
        
        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)
        
        
    Next ws
        


End Sub


'''
