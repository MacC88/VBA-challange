
Sub VBA_Challange()

'Loop through each worksheet
Dim ws As Worksheet
For Each ws In Worksheets
    
    'Set headers for summary table
    ws.Range("I1").Value = ("Ticker")
    ws.Range("J1").Value = ("Yearly Change")
    ws.Range("k1").Value = ("Percent Change")
    ws.Range("l1").Value = ("Total Stock Volume")

    'Set headers for greatest table
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    'Set variable for holding the ticker
    Dim ticker As String
    
    'Set variable for the open values
    Dim open_value As Double
    
    'Keep track of the location for each yearly change
    Dim yearly_Table As Double
    yearly_Table = 2
    
    'Set variable for yearly value
    Dim yearly_value As Double
    
    'Set variable for percent change
    Dim percent_change As Double
    
    'Set variable for the close value
    Dim close_value As Double
    
    'Set variable for holding the total per ticker
    Dim ticker_total As Double
    ticker_total = 0

    'Set variable for the greatest % increase
    Dim percent_increase As Double

    'Set variable for the greatest % increase ticker
    Dim percent_increase_ticker As String

    'Set variable for the greatest % decrease
    Dim percent_decrease As Double

    'Set variable for the greatest % decrease ticker
    Dim percent_decrease_ticker As String
    
    'Set variable for the greatest total
    Dim greatest_total As Double

    'Set variable for the greatest total ticker
    Dim greatest_total_ticker As String

    'Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Define last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop through all tickers
    For i = 2 To lastrow
    
        open_value = ws.Cells(yearly_Table, 3).Value
        
    'Check if still within the same ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Set the ticker
        ticker = ws.Cells(i, 1).Value
        
        'Add to the ticker total
        ticker_total = ticker_total + ws.Cells(i, 7).Value
        
        'Get the close value
        close_value = ws.Cells(i, 6).Value
        
        'Calculate yearly change
        yearly_change = close_value - open_value
        ws.Cells(i, 10).Value = yearly_change
        
            'Calculate percent change
            If open_value = 0 Then
                percent_change = 0
            Else
                percent_change = yearly_change / open_value
            End If
        
        'Print the ticker, yearly change, percent change and ticker total in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = ticker
        ws.Range("J" & Summary_Table_Row).Value = yearly_change
        ws.Range("K" & Summary_Table_Row).Value = Format(percent_change, "Percent")
        ws.Range("L" & Summary_Table_Row).Value = ticker_total
    
            'Format the yearlychange
            If ws.Cells(Summary_Table_Row, 10).Value > 0 Then
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            End If

        'Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Reset the ticker Total
        ticker_total = 0
        
        'Holds the total
        yearly_Table = (i + 1)
        
    'If the cell immediately following a row is the same ticker
    Else

        'Add to the ticker total
        ticker_total = ticker_total + ws.Cells(i, 7).Value
        
    End If
    
  Next i
    
    For i = 2 To lastrow
    
        'Find the ticker with the greatest percent increase
        If ws.Cells(i, 11).Value > percent_increase Then
            percent_increase = ws.Cells(i, 11).Value
            percent_increase_ticker = ws.Cells(i, 9).Value
        
        'Find the ticker with the greatest percent decrease
        ElseIf ws.Cells(i, 11).Value < percent_decrease Then
            percent_decrease = ws.Cells(i, 11).Value
            percent_decrease_ticker = ws.Cells(i, 9).Value
        
        'Find the ticker with the greatest stock volum
        ElseIf ws.Cells(i, 12).Value > greatest_total Then
            greatest_total = ws.Cells(i, 12).Value
            greatest_total_ticker = ws.Cells(i, 9).Value
        End If
    
    Next i
    
   'Add the values for greatest percent increase, decrease and greatest total to each worksheet
    ws.Range("P2").Value = Format(percent_increase_ticker, "Percent")
    ws.Range("Q2").Value = Format(percent_increase, "Percent")
    ws.Range("P3").Value = Format(percent_decrease_ticker, "Percent")
    ws.Range("Q3").Value = Format(percent_decrease, "Percent")
    ws.Range("P4").Value = greatest_total_ticker
    ws.Range("Q4").Value = greatest_total

    'Autofit all
    ws.Range("A:Q").Columns.AutoFit

Next ws

End Sub