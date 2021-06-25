Attribute VB_Name = "Module1"
Sub multiple_year_stockdata()
  
  'Set initial variables
  Dim Opening_Price As Double
  Dim Closing_Price As Double
  Dim Min_Date As Double
  Dim Max_Date As Double
  Dim Yearly_Change As Double
  Dim Percent_Change As String
  Dim Total_Volume As Double
  Dim Ticker_Symbol As String
  Dim Summary_Table_Row As Integer
  Dim ws As Worksheet
  
  Total_Volume = 0
  Opening_Price = 0
  Closing_Price = 0
  Summary_Table_Row = 2

'Loop through all Worksheets
For Each ws In Worksheets

    'Add headers and autofit
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("I1:L1").Columns.AutoFit
    ws.Range("I2:L100000").ClearContents
    ws.Range("I2:L100000").ClearFormats
    Ticker_Symbol = ws.Cells(2, 1)
    
    'Find min and max dates in column B
    Min_Date = WorksheetFunction.MinIfs(ws.Range("B:B"), ws.Range("A:A"), Ticker_Symbol)
    Max_Date = WorksheetFunction.MaxIfs(ws.Range("B:B"), ws.Range("A:A"), Ticker_Symbol)
    
    ' Find the last row of each worksheet
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through ticker column
    For i = 2 To lastrow
     
        ' If same Ticker symbol then...
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
        
                ' Set Ticker Symbol
                Ticker_Symbol = ws.Cells(i, 1)
                
                'Add to the Volume
                Total_Volume = Total_Volume + ws.Cells(i, 7)
                
                ' Print Ticker and Total Volume in the Summary Table
                ws.Range("I" & Summary_Table_Row) = Ticker_Symbol
                ws.Range("L" & Summary_Table_Row) = Total_Volume
                
                'Reset Total_Volume
                Total_Volume = 0
                
                'If MinDate then set opening price...
                Opening_Price = WorksheetFunction.SumIfs(ws.Range("C:C"), ws.Range("A:A"), Ticker_Symbol, ws.Range("B:B"), Min_Date)
        
                ' Otherwise set closing price
                Closing_Price = WorksheetFunction.SumIfs(ws.Range("F:F"), ws.Range("A:A"), Ticker_Symbol, ws.Range("B:B"), Max_Date)
                
            'Set variables
            Yearly_Change = Closing_Price - Opening_Price
            ws.Range("J" & Summary_Table_Row) = Yearly_Change
            
            Ticker_Symbol = ws.Cells(i + 1, 1)
            
            'Reset Min & Max for each ticker_symbol
            Min_Date = WorksheetFunction.MinIfs(ws.Range("B:B"), ws.Range("A:A"), Ticker_Symbol)
            Max_Date = WorksheetFunction.MaxIfs(ws.Range("B:B"), ws.Range("A:A"), Ticker_Symbol)
            
            'Set Percent Change
            If Opening_Price = 0 And Closing_Price = 0 Then
                Percent_Change = 0
            ElseIf Opening_Price = 0 And Closing_Price <> 0 Then
                Percent_Change = 1
            ElseIf Opening_Price <> 0 And Closing_Price = 0 Then
                Percent_Change = -1
            Else
                Percent_Change = FormatPercent(Yearly_Change / Opening_Price, 2)
            End If
            ws.Range("K" & Summary_Table_Row) = Percent_Change
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
            'Conditional formatting to highlight changes
            If Yearly_Change >= 0 Then
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            End If
        
            ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
        Else
        
            'Add to the Total Volume
             Total_Volume = Total_Volume + ws.Cells(i, 7)
        End If
 
    Next i
    
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Total_Volume As Double
    
    'Add headers and autofit
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N1:P1").Columns.AutoFit
    
    'Add rows and autofit
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O:P").Columns.AutoFit
    
    'Find min and max % from column K
    Greatest_Increase = WorksheetFunction.Max(ws.Range("K:K"))
    Greatest_Decrease = WorksheetFunction.Min(ws.Range("K:K"))
    Greatest_Total_Volume = WorksheetFunction.Max(ws.Range("L:L"))
    
    'Print values
    ws.Range("P2") = Greatest_Increase
    ws.Range("P3") = Greatest_Decrease
    ws.Range("P4") = Greatest_Total_Volume
    ws.Range("P2:P3").NumberFormat = "0.00%"
    
    'Find and print associated tickers
    For j = 2 To lastrow
        
        If ws.Cells(j, 11) = Greatest_Increase Then
            ws.Range("O2") = ws.Cells(j, 9)
        End If
    
      
        If ws.Cells(j, 11) = Greatest_Decrease Then
            ws.Range("O3") = ws.Cells(j, 9)
        End If
        
          
        If ws.Cells(j, 12) = Greatest_Total_Volume Then
            ws.Range("O4") = ws.Cells(j, 9)
        End If
        
    Next j
    
Next ws

End Sub
