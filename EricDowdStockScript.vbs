
Sub dowdstockscript()

For Each ws In Worksheets


' Set an initial variable for holding the current stock ticker
  Dim Current_Ticker As String
  Ticker = Cells(2, 1).Value
  
  
' Keep track of stock price at year open for later calculation in summary table
  Dim opening_stockprice As Double
  opening_stockprice = Cells(2, 3).Value
 
 ' Set an initial variable for holding the total stock volume
  Dim Tot_Stock_Volume As Double
  Tot_Stock_Volume = 0
  
' Keep track of the location for each yealy stock summary in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  
' Determine the Last Row
  Dim LastRow As Double
  
  
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Create Stock Summary Chart Headers

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

'Loop through all stock data

    For i = 2 To LastRow
    
    ' Look one row ahead to see if stock ticker is different, if so...
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the stock ticker
            Current_Ticker = Cells(i, 1).Value

      ' Add to the Total stock volume
             Tot_Stock_Volume = Tot_Stock_Volume + Cells(i, 7).Value

      ' Print the stock ticker in the Summary Table
             Range("I" & Summary_Table_Row).Value = Current_Ticker
        
      ' Print yearly change in stock price in the Summary Table
             Range("J" & Summary_Table_Row).Value = ((Cells(i, 6).Value) - opening_stockprice)
        
      ' Color the yearly change Red if negative or green if 0 or greater
             If Range("J" & Summary_Table_Row).Value < 0 Then
                  Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
               Else
                  Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            
             End If
        
      ' Print the Percent change in the summary table
        Range("K" & Summary_Table_Row).Value = (Range("J" & Summary_Table_Row).Value) / opening_stockprice
        Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
        
        
        

      ' Print the total stock volume to the Summary Table
        Range("L" & Summary_Table_Row).Value = Tot_Stock_Volume

      ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total Stock Volume
        Tot_Stock_Volume = 0
      ' Reset the opening stock price
        opening_stockprice = Cells(i + 1, 3)

      
        

    ' If the cell immediately following a row is the same stock...
        Else

      ' Add to the total stock volume
        Tot_Stock_Volume = Tot_Stock_Volume + Cells(i, 7).Value

        End If
    Next i
    
Next ws

End Sub
