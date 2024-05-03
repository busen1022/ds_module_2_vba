Sub Stock_Ticker():

' Run for every sheet
For Each ws In ThisWorkbook.Worksheets


' Set up Leaderboard Header
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Quarterly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("K:K").NumberFormat = "0.00%"

' Set up Biggest Movers
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("Q2,Q3").NumberFormat = "0.00%"
        
        
' Set up variables types
Dim Ticker As String
Dim Quarterly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Vol As LongLong
Total_Stock_Vol = 0 ' Starting Volume Counter
Dim Ticker_Open As Double
Ticker_Open = ws.Cells(2, 3).Value ' Starting Ticker Open
Dim Ticker_Close As Double


' Starting leaderboard row counter
Dim Leader_Row As Integer
Leader_Row = 2

' Find last row for loop with *Xpert Help*
Dim lastrow As LongLong
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

Dim i As LongLong

' Set loop for ticker leaderboard
For i = 2 To lastrow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ' Ticker Column
        Ticker = ws.Cells(i, 1).Value
        ws.Range("I" & Leader_Row).Value = Ticker
        
        ' Quarterly Change
        Ticker_Close = ws.Cells(i, 6).Value
        Quarterly_Change = Ticker_Close - Ticker_Open
        ws.Range("J" & Leader_Row).Value = Quarterly_Change
                   
               
        ' Percent Change
        Percent_Change = Quarterly_Change / Ticker_Open
        ws.Range("K" & Leader_Row).Value = Percent_Change
        
        ' Total Stock Volume
        Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
        ws.Range("L" & Leader_Row).Value = Total_Stock_Vol
        
        Leader_Row = Leader_Row + 1
        Total_Stock_Vol = 0 ' Reset Counter
        Ticker_Open = ws.Cells(i + 1, 3).Value ' Reset Open Price
    
    Else
        Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
             
    End If
    
Next i

' Conditional Format "J"
For k = 2 To lastrow
    
    If ws.Cells(k, 10).Value > 0 Then
        ws.Cells(k, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(k, 10).Value < 0 Then
        ws.Cells(k, 10).Interior.ColorIndex = 3
    End If
    
Next k

' Set loop for Biggest Movers
Dim Max As Double
Max = ws.Cells(2, 11).Value
Dim Min As Double
Min = ws.Cells(2, 11).Value
Dim Max_Vol As LongLong
Max_Vol = ws.Cells(2, 12).Value
Dim Great_Ticker As String
Great_Ticker = ws.Cells(2, 9).Value
Dim Worst_Ticker As String
Worst_Ticker = ws.Cells(2, 9).Value
Dim Vol_Ticker As String
Vol_Ticker = ws.Cells(2, 9).Value

Dim j As LongLong

For j = 2 To lastrow

    ' If for Max (Greatest Increase)
    If ws.Cells(j + 1, 11).Value > Max Then
        Max = ws.Cells(j + 1, 11).Value
        Great_Ticker = ws.Cells(j + 1, 9).Value
        ws.Range("Q2").Value = Max
        ws.Range("P2").Value = Great_Ticker
        
    Else:
        ws.Range("Q2").Value = Max
        ws.Range("P2").Value = Great_Ticker
        
    End If

    ' If for Min (Greatest Decrease)
    If ws.Cells(j + 1, 11).Value < Min Then
        Min = ws.Cells(j + 1, 11).Value
        Worst_Ticker = ws.Cells(j + 1, 9).Value
        ws.Range("Q3").Value = Min
        ws.Range("P3").Value = Worst_Ticker
        
    Else:
        ws.Range("Q3").Value = Min
        ws.Range("P3").Value = Worst_Ticker
        
    End If
    
    ' If for Max Volume (Greatest Total Volume)
    If ws.Cells(j + 1, 12).Value > Max_Vol Then
        Max_Vol = ws.Cells(j + 1, 12).Value
        Vol_Ticker = ws.Cells(j + 1, 9).Value
        ws.Range("Q4").Value = Max_Vol
        ws.Range("P4").Value = Vol_Ticker
        
    Else:
        ws.Range("Q4").Value = Max_Vol
        ws.Range("P4").Value = Vol_Ticker
    
    End If

Next j

' Autofit Leaderboard and Biggest Movers
ws.Columns("I:L").AutoFit
ws.Columns("O:Q").AutoFit


Next ws

End Sub


