VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_analysis()
'class assignment for UofO/edx bootcamp that will analyze stock data.

'will apply the actions to all pages in the active workbook
For Each ws In Worksheets
    ws.Activate

    Dim outputrow_counter As Integer
    outputrow_counter = 2

    Dim volume As Double
    volume = 0

    Dim yearopen As Double
    yearopen = Cells(2, "C").Value

    Dim change As Double

    'heads columns with "Ticker", "Yearly Change", Percent Change" and "Total Stock Volume" respectively
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"

    'finds the number of rows that have data
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'loops through all of the non-empty rows and if the column A ticker is different from the following row, then:
    'puts the ticker, yearly change, percent change and total stock volume in the corresponding output columns

    For i = 2 To RowCount
            
        volume = volume + Cells(i, "G").Value

        If Cells(i + 1, "A").Value <> Cells(i, "A").Value Then
            'display the desired values
            Cells(outputrow_counter, "I").Value = Cells(i, "A").Value
            Cells(outputrow_counter, "J").Value = Cells(i, "F").Value - yearopen
            change = Cells(outputrow_counter, "J").Value
            Cells(outputrow_counter, "K").Value = change / yearopen
            Cells(outputrow_counter, "L").Value = volume
            
            'prepare for the next iteration
            outputrow_counter = outputrow_counter + 1
            yearopen = Cells(i + 1, "C")
            volume = 0
        
        End If
    
    Next i
        
    'apply conditional formatting to the yearly change column
    For j = 2 To RowCount
        If Cells(j, "J").Value > 0 Then
            Cells(j, "J").Interior.ColorIndex = 4
        ElseIf Cells(j, "J").Value < 0 Then
            Cells(j, "J").Interior.ColorIndex = 3
        End If
        
    Next j
        


Next ws

    'adds to the sheet labels for the greatest % increase, greatest % decrease and greatest total volume
    Sheets("A").Cells(2, "O").Value = "Greatest % Increase"
    Sheets("A").Cells(3, "O").Value = "Greatest % Decrease"
    Sheets("A").Cells(4, "O").Value = "Greatest Total Volume"
    Sheets("A").Cells(1, "P").Value = "Ticker"
    Sheets("A").Cells(1, "Q").Value = "Value"
    
    'finds and reports the greatest % increase and greatest % decrease and greatest total volume

Dim greatest_percent_change As Double
Dim greatest_percent_change_ticker As String
Dim greates_percent_decrease As Double
Dim greatest_percent_decrease_ticker As String

greatest_percent_change = 0

For Each ws In Worksheets
        
        RowCount_tickertotals = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
    For K = 2 To RowCount_tickertotals
        If ws.Cells(K, "K").Value > greatest_percent_change Then
            greatest_percent_change = ws.Cells(K, "K").Value
            greatest_percent_change_ticker = ws.Cells(K, "I").Value
        End If
        
        If ws.Cells(K, "K").Value < Sheets(1).Cells(3, "Q").Value Then
            Sheets(1).Cells(3, "Q").Value = ws.Cells(K, "K").Value
            Sheets(1).Cells(3, "P").Value = ws.Cells(K, "I").Value
        End If
        
        If ws.Cells(K, "L").Value > Sheets(1).Cells(4, "Q").Value Then
            Sheets(1).Cells(4, "Q").Value = ws.Cells(K, "L").Value
            Sheets(1).Cells(4, "P").Value = ws.Cells(K, "I").Value
        End If
    
    Next K
    
Next ws
Sheets(1).Cells(2, "Q").Value = greatest_percent_change
Sheets(1).Cells(2, "P").Value = greatest_percent_change_ticker

End Sub
