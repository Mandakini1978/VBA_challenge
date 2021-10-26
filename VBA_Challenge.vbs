'This sub routine generates a summary table containg following information :
'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock and 
'conditional formatting on Yearly Change column that will highlight positive change  in green and negative change in red.



Sub VBA_Challenge()

 'looping for each worksheet
 For Each ws In Worksheets

  ' Set an initial variable for holding Ticker Name
  Dim Ticker_Name As String

  ' Set an initial variable for holding opening stock
  Dim Year_Open As Double
  ' Set an initial variable for holding closing stock
  Dim Year_Close As Double
  'counting no of entries for a particular ticker
  Dim Count As Long
  'Set an initial variable for holding stock volume
  Dim Vol As Double
  
   'initializing variables by zero
  Year_Open = 0
  Year_Close = 0
  Count = 0
  Vol = 0

  ' Keep track of the row count of the  summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  'Finding total rows
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all ticker value
  For i = 2 To LastRow

    ' Check if we are still within the same ticker value, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker value
      Ticker_Name = ws.Cells(i, 1).Value

       
      
      Year_Close = ws.Cells(i, 6).Value
      
      Vol = Vol + ws.Cells(i, 7).Value

      'set ticker name to summary table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
        
      'Calculating yearly change to the Summary Table
      
      ws.Range("J" & Summary_Table_Row).Value = Year_Close - Year_Open
      
      'calculating percentage change to the Summary Table
      
      If Year_Open <> 0 Then
            ws.Range("K" & Summary_Table_Row).Value = (Year_Close - Year_Open) / Year_Open
      Else
            ws.Range("K" & Summary_Table_Row).Value = Null
      End If
           
      'calculating Total stock volume  to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Vol
      
      'Formatting row color
      
       If (Year_Close - Year_Open) > 0 Then
            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
       Else
            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
       End If
            
      
      ' Add one to the summary table row
      
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset variables
      Year_Open = 0
      Year_Close = 0
      Count = 0
      Vol = 0

    ' If the cell immediately following a row is the same ticker value...
      Else

      
      Count = Count + 1
      
      Vol = Vol + ws.Cells(i, 7).Value
      
      If Count = 1 Then
        Year_Open = ws.Cells(i, 3).Value
      Else
        Year_Close = ws.Cells(i, 6).Value

        End If
    End If
    
   
  Next i
  'naming columns in the summary table
  
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "YearlyChange"
  ws.Cells(1, 11).Value = "PercentChange"
  ws.Cells(1, 12).Value = "TotalStockVolume"
  
  
  Next ws

  End Sub



Sub Calculate()

'calculating  Greatest percentage Increase in Value, Greatest percentage Decrease in Value
'and Greatest Total Volume with ticker name from summary table


  Dim Max_Val As Double
  Dim Min_Val As Double
  Dim Max_Vol As Double
  
For Each ws In Worksheets

  
  
  Max_Val = ws.Cells(2, 11).Value
  Min_Val = ws.Cells(2, 11).Value
  Max_Vol = ws.Cells(2, 12).Value
  
  
  
  For i = 2 To 290
  
    If Max_Val < ws.Cells(i + 1, 11).Value Then
    
        Max_Val = ws.Cells(i + 1, 11).Value
        Row_Max = i
              
    End If
  Next i
  
  ws.Cells(2, 17).Value = Max_Val
  ws.Cells(2, 16).Value = ws.Cells(Row_Max + 1, 9)
  
    
    For i = 2 To 290
    If Min_Val > ws.Cells(i + 1, 11).Value Then
    
        Min_Val = ws.Cells(i + 1, 11).Value
        Row_Min = i
              
    End If
    Next i
    
    ws.Cells(3, 17).Value = Min_Val
    ws.Cells(3, 16).Value = ws.Cells(Row_Min + 1, 9)
    
    For i = 2 To 290
    If Max_Vol < ws.Cells(i + 1, 12).Value Then
    
        Max_Val = ws.Cells(i + 1, 12).Value
        Row_Max_Vol = i
              
    End If
    Next i
  
    ws.Cells(4, 17).Value = Max_Vol
    ws.Cells(4, 16).Value = ws.Cells(Row_Max_Vol + 1, 9)
        
       
        
    
        
        
     ws.Cells(2, 15).Value = "Greatest % Increase in Value"
     ws.Cells(3, 15).Value = "Greatest % Decrease in Value"
     ws.Cells(4, 15).Value = "Greatest Total Volume"
        
     ws.Cells(1, 16).Value = "Ticker"
     ws.Cells(1, 17).Value = "Value"

        
 Next ws
  
  
End Sub



