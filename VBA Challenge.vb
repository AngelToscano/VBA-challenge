'Create a script that loops through all the stocks for one year and outputs the following information:

'The ticker symbol

'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

'The total stock volume of the stock.


Sub stock_ticker()

  ' Set an initial variable for holding the ticker name, total volume, and Summary table
  Dim ticker As String
  Dim Volume_Total As Double
  Volume_Total = 0
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Determine the number of rows
  RowCount = Cells(Rows.Count, "A").End(xlUp).Row
  
  'Declare Opening Ticker Amount, Closing Ticker Amount, Yearly Change, and Percentage Change
  Dim Open_Ticker As Double
  Dim Closing_Ticker As Double
  Dim Yearly_Change As Double
  Dim Percent_Change As Double
    
  'Lable the Summary Column
  Cells(1, 11).Value = "Ticker"
  Cells(1, 12).Value = "Yearly Change"
  Cells(1, 13).Value = "Percentage Change"
  Cells(1, 14).Value = "Total Stock Volume"
  
  
  ' Loop through all tickers
  For i = 2 To RowCount
  
    ' Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker name
      ticker = Cells(i, 1).Value

      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value

      ' Print the Ticker name in the Summary Table
      Range("K" & Summary_Table_Row).Value = ticker

      ' Print the Volume Amount to the Summary Table
      Range("N" & Summary_Table_Row).Value = Volume_Total
      
      'Set the Closing ticker amount
      Closing_Ticker = Cells(i, 6).Value
       
      'Determine the yearly change and print it in the Summary table and have the cell turn red if negative and green if positive
      Yearly_Change = Closing_Ticker - Open_Ticker
      Range("L" & Summary_Table_Row).Value = Yearly_Change
      If Range("L" & Summary_Table_Row).Value < 0 Then Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
      If Range("L" & Summary_Table_Row).Value > 0 Then Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
        
      'Determine the Percentage change and print it in Summary box as a percentage
      Percent_Change = (Yearly_Change / Open_Ticker)
      Range("M" & Summary_Table_Row).Value = Percent_Change
      Range("M" & Summary_Table_Row).Value = FormatPercent(Range("M" & Summary_Table_Row))
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Volume Total
      Volume_Total = 0

    'If the cell immediately following a row is the same ticker name...
    ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    
    'Set the Open Ticker and Determine the Yearly Change
    Open_Ticker = Cells(i, 3).Value
    Yearly_Change = Open_Ticker - Closing_Ticker
    
    Else

    ' Add to the Volume Total
    Volume_Total = Volume_Total + Cells(i, 7).Value
      
    End If


  Next i

End Sub


