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
    RowCount = Cells(Rows.Count, 1).End(xlUp).Row

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
        
    'applies formatting to the yearly percent column and greatest and least percent results(written by recording a macro then editing)
    Columns("K:K").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Range("Q2").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Range("Q3").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"

    'finds and reports the greatest % increase and greatest % decrease and greatest total volume
    'declares variables
    Dim greatest_percent_increase As Double
    Dim greatest_percent_increase_ticker As String
    Dim greatest_percent_decrease As Double
    Dim greatest_percent_decrease_ticker As String
    Dim greatest_total_volume As Double
    Dim greatest_total_volume_ticker As String

    'adds to the sheet labels for the greatest % increase, greatest % decrease and greatest total volume
    Cells(2, "O").Value = "Greatest % Increase"
    Cells(3, "O").Value = "Greatest % Decrease"
    Cells(4, "O").Value = "Greatest Total Volume"
    Cells(1, "P").Value = "Ticker"
    Cells(1, "Q").Value = "Value"
    
    'initializes variables for this sheet
    RowCount_tickertotals = Cells(Rows.Count, 9).End(xlUp).Row
    greatest_percent_increase = 0
    greatest_percent_decrease = 0
    greatest_total_volume = 0
    
    For K = 2 To RowCount_tickertotals
        If Cells(K, "K").Value > greatest_percent_increase Then
            greatest_percent_increase = Cells(K, "K").Value
            greatest_percent_increase_ticker = Cells(K, "I").Value
        End If
        
        If Cells(K, "K").Value < greatest_percent_decrease Then
            greatest_percent_decrease = Cells(K, "K").Value
            greatest_percent_decrease_ticker = Cells(K, "I").Value
        End If
        
        If Cells(K, "L").Value > greatest_total_volume Then
            greatest_total_volume = Cells(K, "L").Value
            greatest_total_volume_ticker = Cells(K, "I").Value
        End If
    
    Next K
    
    Cells(2, "Q").Value = greatest_percent_increase
    Cells(2, "P").Value = greatest_percent_increase_ticker
    Cells(3, "Q").Value = greatest_percent_decrease
    Cells(3, "P").Value = greatest_percent_decrease_ticker
    Cells(4, "Q").Value = greatest_total_volume
    Cells(4, "P").Value = greatest_total_volume_ticker
Next ws

End Sub

