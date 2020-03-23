Attribute VB_Name = "Module1"
Sub Ticker_Symbol()
   
    'Create a variable for worksheet
    Dim ws As Worksheet
    
    'Loop through all sheets
    For Each ws In Worksheets
        ws.Activate
        Debug.Print ws.Name
    
    'Create a varaible for the lastrow
    Dim LastRow As Double
    
    'Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(x1Up).Row
   
    'Create variable for tickersymbol
    Dim Tickersymbol As String
    
    ' Set an initial variable for holding the total volume per stock
    Dim StockVolume As LongLong
    StockVolume = 0
    
    'Create variable for year open, year close, yearly change, percent change
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
     
    'Keep track of the location for each ticker symbol in the ticker table
    Dim Ticker_Table_Row As Integer
    Ticker_Table_Row = 1
        
        'Create headers in alphabetical_testing.Worksheets
        ws.Cells(1, 9).Value = "Tickersymbol"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "StockVolume"

     'Loop through all tickersymbol
        For i = 2 To LastRow
    
          'Check if we are still within the same tickersymbol, if it is not..
          If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          
          'Set Tickersymbol name
          Tickersymbol = Cells(i, 1).Value
          
          'Add to the StockVolume
          StockVolume = StockVolume = Cells(i, 7)
          
          'Add to the Ticker Table Row
           Ticker_Table_Row = Tickerysymbol + ws.Cells(i, 9).Value
            
              'Print the tickersymbol in the tickersymbol to Ticker Table
              ws.Range("I" & Ticker_Table_Row).Value = Tickersymbol
            
              'Print the Stock total to the Ticker Table
              ws.Range("L" & Ticker_Table_Row).Value = StockVolume
            
              'Add one to the ticker table row table
              Ticker_Table_Row = Ticker_Table_Row + 1
            
              'Reset the Stock Total
              StockVolume = 0
             
             'If the cell immdeiately following a row is the same ticker symbol...
              Else
        
               'Add to the StockVolume Total
                StockVolume = StockVolume + ws.Cells(i, 7).Value
            
            'End If
            End If
      
        'End loop
        Next i
        
    'next worksheet
    Next ws
      
End Sub
               
