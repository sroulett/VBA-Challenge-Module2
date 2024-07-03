# VBA-Challenge-Module2

Sub Ticker_Challenge()

Dim ticker As String
Dim open_price As Double
Dim closing_price As Double

Dim price As Double
Dim ttlsv As Double
Dim yrc As Double
Dim gtv As Double

Dim PreviousStockPrice As Long
Dim table_summary_row As Long
Dim greatest_increase As Double
Dim greatest_decrease As Double

Dim ws As Worksheet
For Each ws In Worksheets

'Headers and Tables
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ttlsv = 0
table_summary_row = 2
PreviousStockPrice = 2

LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRowA

        ttlsv = ttlsv + ws.Cells(i, 7).Value

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                ws.Range("I" & table_summary_row).Value = ticker
                ws.Range("L" & table_summary_row).Value = ttlsv
                ttlsv = 0

                'Year open and close price, yearly change, and percentage change
                open_price = ws.Range("C" & PreviousStockPrice)
                close_price = ws.Range("F" & i)
                yrc = close_price - open_price
                ws.Range("J" & table_summary_row).Value = yrc

                If open_price = 0 Then
                    price = 0

                Else
                    open_pice = ws.Range("C" & PreviousStockPrice)
                    price = yrc / open_price

                End If

                
                ws.Range("K" & table_summary_row).Value = price
                ws.Range("K" & table_summary_row).NumberFormat = "0.00%"
                If ws.Range("J" & table_summary_row).Value >= 0 Then
                    ws.Range("J" & table_summary_row).Interior.ColorIndex = 4

                Else

  ws.Range("J" & table_summary_row).Interior.ColorIndex = 3

                End If
                table_summary_row = table_summary_row + 1
                PreviousStockPrice = i + 1

            End If

            Next i

greatest_increase = 0
greatest_decrease = 0
gtv = 0

LastRowK = ws.Cells(Rows.Count, 11).End(xlUp).Row

For i = 2 To LastRowK

    'Greatest % Increase
    If ws.Range("K" & i).Value > greatest_increase Then
        greatest_increase = ws.Range("K" & i).Value
        ws.Range("Q2").Value = greatest_increase
        ws.Range("P2").Value = ws.Range("I" & i).Value

    End If

    'Greatest % Decrease
    If ws.Range("K" & i).Value < greatest_decrease Then
        greatest_decrease = ws.Range("K" & i).Value
        ws.Range("Q3").Value = greatest_decrease
        ws.Range("P3").Value = ws.Range("I" & i).Value

    End If

 'Greatest Total Volume
    If ws.Range("L" & i).Value > gtv Then
       gtv = ws.Range("L" & i).Value
       ws.Range("Q4").Value = gtv
       ws.Range("P4").Value = ws.Range("I" & i).Value

    End If

 'Format to "%" for Greatest % Increase and Decrease
    ws.Range("Q2").NumberFormat = "0.00%"

    ws.Range("Q3").NumberFormat = "0.00%"

Next i

Next ws


End Sub


Citations:
https://stackoverflow.com/questions/77118851/how-do-i-get-the-largest-percent-change-with-the-corresponding-ticker-using-if-s

https://stackoverflow.com/questions/62471422/vba-loop-how-to-get-ticker-symbols-into-ticker-column

Tutoring session with Brandon W.
