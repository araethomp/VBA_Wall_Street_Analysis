Attribute VB_Name = "Module6"
Sub VBA_Wall_Street_Analysis()

    Dim WS As Worksheet

    For Each WS In Worksheets
        WS.Activate
        
        
        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 10).Value = "Yearly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Stock Volume"


        WS.Cells(1, 16).Value = "Ticker"
        WS.Cells(1, 17).Value = "Value"
        WS.Cells(2, 15).Value = "Greatest % Increase"
        WS.Cells(3, 15).Value = "Greatest % Decrease"
        WS.Cells(4, 15).Value = "Greatest Total Volume"


        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).row


        'Variables
        Dim ticker_name As String
        Dim ticker_total As Double
        Dim diff As Double
        Dim row As Integer
        Dim i As Long
        Dim j As Long

        'Initialize Variables
        ticker_total = 0
        j = 2
        diff = 0
        ticker = 2
        
        For i = 2 To LastRow


            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                ticker_name = Cells(i, 1).Value
                 ticker_total = ticker_total + Cells(i, 7).Value

                If ticker_total = 0 Then
                    Cells(ticker, 9).Value = Cells(i, 1).Value
                    Cells(ticker, 10).Value = 0
                    Cells(ticker, 11).Value = 0 & "%"
                    Cells(ticker, 12).Value = 0
                Else
                    If Cells(j, 3) = 0 Then
                    
                        For pointer = j To i
                            If Cells(pointer, 3).Value <> 0 Then
                                j = pointer
                                Exit For
                            End If
                        Next pointer
                    End If
                    
                
                    diff = (Cells(i, 6).Value - Cells(j, 3).Value)
                    
                   
                    percent_change = Round((diff / Cells(j, 3) * 100), 4)

                    
                    Cells(ticker, 9).Value = ticker_name
                    Cells(ticker, 10).Value = diff
                    Cells(ticker, 11).Value = "%" & percent_change
                    Cells(ticker, 12).Value = ticker_total

                    If diff > 0 Then
                            Cells(ticker, 10).Interior.ColorIndex = 4
                    ElseIf diff < 0 Then
                            Cells(ticker, 10).Interior.ColorIndex = 3
                    Else
                            Cells(ticker, 10).Interior.ColorIndex = 2
                    End If
                End If
 
                    ticker = ticker + 1
                    j = i + 1
                    ticker_total = 0
                    diff = 0

           Else
                ticker_total = ticker_total + Cells(i, 7).Value
           End If

        Next i
        
        'Bonus Challenges
        greatest_percent_increase = WorksheetFunction.Max(Range("K2:K" & LastRow)) * 100
        Range("Q2") = greatest_percent_increase & "%"
     
        Position = WorksheetFunction.Match(Range("Q2").Value, Range("K2:K" & LastRow), 0)
        Range("P2") = Cells(Position + 1, 9)
        
        greatest_percent_decrease = WorksheetFunction.Min(Range("K2:K" & LastRow)) * 100
        Range("Q3") = greatest_percent_decrease & "%"
        
        Position = WorksheetFunction.Match(Range("Q3").Value, Range("K2:K" & LastRow), 0)
        Range("P3") = Cells(Position + 1, 9)
        
        greatest_total_volume = WorksheetFunction.Max(Range("L2:L" & LastRow))
        Range("Q4") = greatest_total_volume
        
        Position = WorksheetFunction.Match(Range("Q4").Value, Range("L2:L" & LastRow), 0)
        Range("P4") = Cells(Position + 1, 9)

    Next WS
End Sub

'Citation:
'Title: VBA_Reference
'Author: Chao, D.
'Date: 2020
'Code Version: 1.0
'Availability: Used as reference while working on homework during tutoring sessions

