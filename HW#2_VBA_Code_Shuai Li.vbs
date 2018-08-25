Sub Stock_market_analyst()

  For Each ws in Worksheets

  ' Set an initial variable for holding the ticker symbol
  Dim Ticker_symbol As String

  ' Set an initial variable for holding the total stock volume per ticker
  Dim Total_Stock_Volume As Double
  Total_Stock_Volume = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  ' Create a variable to hold the ticker Counter
  Dim Ticker_Counter As Integer 
  Ticker_Counter = 1

  ' Set an initial variable for holding the yearly change
  Dim Yearly_Change As Double

  ' Set an initial variable for holding the percent change
  Dim Percent_Change As Double

  Dim Greatest_Icrease As Double
  Dim Greatest_Decrease As Double
  Dim Greatest_Volume As Double

  Dim Greatest_Icrease_Row As Integer
  Dim Greatest_Decrease_Row As Integer
  Dim Greatest_Volume_Row As Integer

  ' Add the Name to the First Row Header
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Total Stock volume"
  ws.Range("K1").Value = "Yearly Change"
  ws.Range("L1").Value = "Percent Change"
  
  ws.Range("O1").Value = "Ticker"
  ws.Range("P1").Value = "Value"
  ws.Cells(2, 14).Value = "Greatest % Increase"
  ws.Cells(3, 14).Value = "Greatest % Decrease"
  ws.Cells(4, 14).Value = "Greatest Total Volume"

  ' Determine the Last Row
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all ticker's volume everyday
  For i = 2 To LastRow 

    ' Check if we are still within the same ticker, if it is the same ticker... 
    If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then

      ' Add to the total stock volume
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

      ' Add 1 to the ticker counter
      Ticker_Counter = Ticker_Counter + 1

    ' If it is not...
    Else

      ' Set the ticker symbol
      Ticker_symbol = ws.Cells(i, 1).Value

      ' Print the ticker symbol in the summary table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_symbol

      ' Print the total stock volume to the summary table
      ws.Range("J" & Summary_Table_Row).Value = Total_Stock_Volume

      ' Set the yearly change
      Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(i - Ticker_Counter + 1, 3).Value  

      ' Print the yearly change to the summary table
      ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
      
      If ws.Cells(i - Ticker_Counter + 1, 3).Value = 0 Then
      Percent_Change = 0
      Else 
      Percent_Change = Yearly_Change / ws.Cells(i - Ticker_Counter + 1, 3).Value 
      End if 

      ' Print the percent change to the summary table
      ws.Range("L" & Summary_Table_Row).Value = Percent_Change

      'Range("L" & Summary_Table_Row).Style = "Percentage"

      ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%" 

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

      ' Reset the total stock volume
      Total_Stock_Volume = 0

      ' Reset the ticker counter volume
      Ticker_Counter = 1
    
    End if 
    
     ' USe if statment to avoid denominator = 0 error
    If ws.Range("K" & Summary_Table_Row).Value < 0 Then
       ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
    Else
       ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
    End if
    
  Next i

  ' Determine the Last Row of Table
  LastRow_Table = ws.Cells(Rows.Count, 9).End(xlUp).Row

  'Output 
  Greatest_Icrease = WorksheetFunction.Max(ws.Range("L2:L"&LastRow_Table))
  ws.Cells(2, 16).Value = Greatest_Icrease
  ws.Cells(2, 16).NumberFormat = "0.00%"
  Greatest_Icrease_Row = WorksheetFunction.Match(Greatest_Icrease,ws.Range("L1:L"&LastRow_Table),0)
  ws.Cells(2, 15).Value = ws.Cells(Greatest_Icrease_Row, 9).Value 

  Greatest_Decrease = WorksheetFunction.Min(ws.Range("L2:L"&LastRow_Table))
  ws.Cells(3, 16).Value = Greatest_Decrease
  ws.Cells(3, 16).NumberFormat = "0.00%"
  Greatest_Decrease_Row = WorksheetFunction.Match(Greatest_Decrease,ws.Range("L1:L"&LastRow_Table),0)
  ws.Cells(3, 15).Value = ws.Cells(Greatest_Decrease_Row, 9).Value

  Greatest_Volume = WorksheetFunction.Max(ws.Range("J2:L"&LastRow_Table))
  ws.Cells(4, 16).Value = Greatest_Volume
  Greatest_Volume_Row = WorksheetFunction.Match(Greatest_Volume,ws.Range("J1:J"&LastRow_Table),0)
  ws.Cells(4, 15).Value = ws.Cells(Greatest_Volume_Row, 9).Value

  ' Autofit to display data
  ws.Columns("A:P").AutoFit

  Next ws 

End Sub

