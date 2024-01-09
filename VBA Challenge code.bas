Attribute VB_Name = "Module1"
Sub stockData():
    
    Dim total As Double
    Dim i As Long
    Dim j As Integer
    Dim change As Double
    Dim start As Long
    Dim rowCount As Long
    Dim percentChanged As Double
    Dim ws As Worksheet
    Dim k As Long
    
    For Each ws In Worksheets
        
        j = 0
        change = 0
        total = 0
        start = 2
        
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"
        ws.Cells(1, "R").Value = "Ticker"
        ws.Cells(1, "S").Value = "Value"
        ws.Cells(2, "Q").Value = "Greatest Percent Increase"
        ws.Cells(3, "Q").Value = "Greatest Percent Decrease"
        ws.Cells(4, "Q").Value = "Greatest Total Volume"
        
        ws.Columns("A:S").AutoFit
        
        For k = 1 To 1000000
            
            If (IsEmpty(Cells(k, 1)) = False) Then
            
                rowCount = rowCount + 1
                
            End If
            
        Next k
        

        For i = 2 To rowCount
            
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                
                total = total + ws.Cells(i, 7).Value
                
                If (total = 0) Then
                
                    ws.Range("I" & (2 + j)).Value = Cells(i, 1).Value
                    ws.Range("J" & (2 + j)).Value = 0
                    ws.Range("K" & (2 + j)).Value = "%" & 0
                    ws.Range("L" & (2 + j)).Value = 0
                    
                Else
                
                    If (ws.Cells(start, 3) = 0) Then
                        
                        For findValue = start To i
                            
                            If (ws.Cells(findValue, 3).Value <> 0) Then
                            
                                start = findValue
                                Exit For
                            
                            End If
                            
                        Next findValue
                        
                    End If
                
                change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                percentChanged = (change / ws.Cells(start, 3))
                
                start = i + 1
                
                ws.Range("I" & (2 + j)).Value = ws.Cells(i, 1).Value
                ws.Range("J" & (2 + j)).Value = change
                ws.Range("J" & (2 + j)).NumberFormat = "0.00"
                ws.Range("K" & (2 + j)).Value = percentChanged
                ws.Range("K" & (2 + j)).NumberFormat = "0.00%"
                ws.Range("L" & (2 + j)).Value = total
                
                If (ws.Cells(2 + j, "J").Value > 0) Then
                    
                    ws.Cells(2 + j, "J").Interior.ColorIndex = 4
                    
                ElseIf (ws.Cells(2 + j, "J").Value < 0) Then
                    
                    ws.Cells(2 + j, "J").Interior.ColorIndex = 3
                    
                Else
                    
                    ws.Cells(2 + j, "J").Interior.ColorIndex = 0
                
                End If
                
                
            End If
            
            total = 0
            change = 0
            j = j + 1
                
            Else
            
                total = total + ws.Cells(i, 7).Value
            
            End If
            
          Next i
          
          ws.Cells(2, "S") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
          ws.Cells(3, "S") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
          ws.Cells(4, "S") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
          
          ws.Range("S2").NumberFormat = "0.00%"
          ws.Range("S3").NumberFormat = "0.00%"
          ws.Range("S4").NumberFormat = "0.00"
          
          increaseNum = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
          decreaseNum = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
          volumeNum = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
          
          ws.Cells(2, "R") = ws.Cells(increaseNum + 1, 9)
          ws.Cells(3, "R") = ws.Cells(decreaseNum + 1, 9)
          ws.Cells(4, "R") = ws.Cells(volumeNum + 1, 9)
        
        Next ws
    
End Sub
