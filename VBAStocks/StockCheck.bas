Attribute VB_Name = "Module1"
Sub StockCheck()

Dim WS, startWS As Worksheet
Dim Ticker, maxticker, minticker As String
Dim OpenPrice, ClosePrice, PercentChange, Volume, SummaryRow, pmax, pmin As Double
Dim cfrg, prg As Range

Set startWS = ActiveSheet

'Loop through all worksheets
For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increased"
    Range("O3").Value = "Greatest % Decreased"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    pmax = 0
    pmin = 0
    SummaryRow = 2
    OpenPrice = Cells(2, 3).Value
    Volume = 0
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        'If reach the end of a ticker, record it in the summary table
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            ClosePrice = Cells(i, 6).Value
            Volume = Volume + Cells(i, 7).Value
            'Conditional formating yearly price difference with different colors
            Set cfrg = Range("J" & SummaryRow)
            cfrg.Value = ClosePrice - OpenPrice
            Select Case cfrg
                Case Is >= 0
                    With cfrg
                        .Interior.Color = vbGreen
                    End With
                Case Is < 0
                    With cfrg
                        .Interior.Color = vbRed
                    End With
            End Select
            'There are cases the open price = 0 which will cause overflow
            If OpenPrice <> 0 Then
                PercentChange = FormatPercent((ClosePrice - OpenPrice) / OpenPrice)
            Else
                PercentChange = 0
            End If
            
            Range("I" & SummaryRow).Value = Ticker
            Range("K" & SummaryRow).Value = PercentChange
            Range("L" & SummaryRow).Value = Volume
         
            'Check for max volume
            If Volume > Range("Q4").Value Then
                Range("Q4").Value = Volume
                Range("P4").Value = Ticker
            End If
            
            'Reset open price and volume for next ticker
            OpenPrice = Cells(i + 1, 3).Value
            Volume = 0
            SummaryRow = SummaryRow + 1
        'If the next row is the same ticker, keep adding up total volume
        Else
            If OpenPrice = 0 Then
                OpenPrice = Cells(i + 1, 3).Value
            End If
            Volume = Volume + Cells(i, 7).Value
        End If
    Next i
    
    'Check for max and min percent change
    Set prg = Range("k2", Range("k2").End(xlDown))
    For j = 2 To prg.Count
        If Cells(j, "K").Value > pmax Then
            pmax = Cells(j, "K").Value
            maxticker = Cells(j, "I").Value
        ElseIf Cells(j, "K").Value < pmin Then
            pmin = Cells(j, "K").Value
            minticker = Cells(j, "I").Value
        End If
    Next j
    
    Range("P2").Value = maxticker
    Range("Q2").Value = FormatPercent(pmax)
    Range("P3").Value = minticker
    Range("Q3").Value = FormatPercent(pmin)
    
Next

End Sub
