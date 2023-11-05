' =========================================================================================== '
' Author: Hamza Saleem
' Instructor: Piro Dhimitri
' =========================================================================================== '

Sub stocks():

    ' Loop through each worksheet
    For Each ws In Worksheets
        
            ' Create header for Ticker, Year change, percent Change, Stock Volume
                ws.Range("J1").Value = "Ticker"
                ws.Range("K1").Value = "Year Change"
                ws.Range("L1").Value = "Percent Change"
                ws.Range("M1").Value = "Total Stock Volume"
                
            ' Set the total number of rows
            Dim rowNum As Long
            rowNum = ws.Range("A1").End(xlDown).Row
            
            ' Set an initial variable for holding the Ticker
            Dim ticker As String

            ' Set an initial variable for holding the total stock volume
            Dim stockVolume As Double
            stockVolume = 0
            
            ' Set an initial variable for holding Year Change, opening balance and closing Balance
            Dim opening, closing, yrChange As Double
            yrChange = 0
            opening = 0
            closing = 0
            
            ' Set a flag to determine if the entry is the first entry in the ticker.
            ' This will help determine the openeing balance of the stock
            openingFlag = 1
            
            ' Set an initial variable for holding percentage change
            Dim percentChange As Double
            percentChange = 0
        
            ' Keep track of rows for each new ticker in the summary table
            Dim summaryTableRow As Long
            summaryTableRow = 2
            
            ' Loop through all stocks
             For i = 2 To rowNum

' Check if we are still within the same ticker, if it is not...
                If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                    ' --------------  Ticker ------------- '
                  ' Set the Ticker
                    ticker = ws.Cells(i, 1).Value
                    
                  ' Print the ticker in the Summary Table
                    ws.Range("J" & summaryTableRow).Value = ticker
            
                    ' --------------- Total Stock Volume ----------- '
                  ' Add to the stockVolume
                    stockVolume = stockVolume + ws.Cells(i, 7).Value
 
                  ' Print the stock volume to the Summary Table
                    ws.Range("M" & summaryTableRow).Value = stockVolume
                    
                   ' Reset the Stock Volume Total
                    stockVolume = 0
                    
                    ' ------------------ Year Change --------------- '
                  ' set the closing balance
                    closing = ws.Cells(i, 6).Value
                   
                  ' Print the year change to the Summary Table
                    yrChange = closing - opening
                    ws.Range("K" & summaryTableRow).Value = yrChange
                    
                  ' Conditional Formatting for coloring year changes
                  
                    ' If the change is higher than last years then color green
                    If (yrChange > 0) Then
                        ws.Range("K" & summaryTableRow).Interior.ColorIndex = 4
                        
                    ' If the change is lower than last years then color red
                    ElseIf (yrChange < 0) Then
                        ws.Range("K" & summaryTableRow).Interior.ColorIndex = 3
                        
                    ' If there is no change then color blue
                    Else
                        ws.Range("K" & summaryTableRow).Interior.ColorIndex = 8
                        
                    End If

                  ' Set the openingFlag openingFlag to true
                    openingFlag = 1
                    
                    ' ------------------ Percentage Change ---------------- '
                  ' Percent change would be yearly change divided by the opening balance as a percentage
                    percentChange = (yrChange / opening)
                    ws.Range("L" & summaryTableRow).Value = percentChange
                    ws.Range("L" & summaryTableRow) = Format(percentChange, "0.00%")
                    
                  ' Conditional Formatting for coloring percent changes
                  
                    ' If the change is higher than last years' then color green
                    If (percentChange > 0) Then
                        ws.Range("L" & summaryTableRow).Interior.ColorIndex = 4
                        
                    ' If the change is lower than last years then color red
                    ElseIf (percentChange < 0) Then
                        ws.Range("L" & summaryTableRow).Interior.ColorIndex = 3
                        
                    ' If there is no change then color blue
                    Else
                        ws.Range("L" & summaryTableRow).Interior.ColorIndex = 8
                        
                    End If
                   
                    ' ----------- creating new row in summary table ------ '
                ' Add one to the summary table row
                    summaryTableRow = summaryTableRow + 1
                    
' If the cell immediately following a row is the same ticker...
                Else
                     ' --------------- Total Stock Volume ----------- '
                  ' Add to the Stock Volume Total
                    stockVolume = stockVolume + ws.Cells(i, 7).Value
                    
                        ' ------------------ Year Change --------------- '
                    ' If the openining flag is true, this is the first entry of the ticker ie the opening balance
                      If (openingFlag = 1) Then
                        ' set the opening value of the ticker
                        opening = ws.Cells(i, 3).Value
                        
                        ' set the first value flag to false
                        openingFlag = 0
                      End If
                 
                End If
               
          Next i
          
          ' ====================================================================================== '
          ' ---- Table for greatest % increase, Greatest % decrease, and Greatest Stock Volume --- '
          ' ====================================================================================== '
          
          ' Create headers for new table
            ws.Range("P2").Value = "Greatest % Increase"
            ws.Range("P3").Value = "Greatest % Decrease"
            ws.Range("P4").Value = "Greatest Total Volume"
            ws.Range("Q1").Value = "Ticker"
            ws.Range("R1").Value = "Value"
          
          ' Set the row count for the new table
            Dim rowNumTickers As Long
            rowNumTickers = ws.Range("J1").End(xlDown).Row
            
          ' Initialize and set the first values in the table
          Dim greatestPercentIncrease As Double
          greatestPercentIncrease = 0
          'greatestPercentIncrease = ws.Range("L2").Value
          
          Dim greatestPercentDecrease As Double
          greatestPercentDecrease = 0
          'greatestPercentDecrease = ws.Range("L2").Value
          
          Dim greatestStockVolume As Double
          greatestStockVolume = 0
          'greatestStockVolume = ws.Range("M2").Value
          
          ' Initialize variables to hold the tickers
          Dim tickerIncrease As String
          Dim tickerDecrease As String
          Dim tickerVolume As String
         
          ' For loop to go through all the tickers
            For x = 2 To rowNumTickers
            ' Check to see if the current value is bigger than the next one
                If (greatestPercentIncrease < ws.Range("L" & x).Value) Then
                ' If the current value is smaller than the value in the cell then assign the value of the cell as greatest percent increase
                    greatestPercentIncrease = ws.Range("L" & x).Value
                    tickerIncrease = ws.Range("J" & x).Value
                End If
                
            ' Check to see if the current value is smaller than the one in the cell
                If (greatestPercentDecrease > ws.Range("L" & x).Value) Then
                ' If the current value is larger than the value in the cell then assign the value of the cell as greatest percent decrease
                    greatestPercentDecrease = ws.Range("L" & x).Value
                    tickerDecrease = ws.Range("J" & x).Value
                End If
            
            ' Check to see if the current value is smaller than the one in the cell
                If (greatestStockVolume < ws.Range("M" & x).Value) Then
                ' If the current value is larger smaller the value in the cell then assign the value of the cell as greatest percent decrease
                    greatestStockVolume = ws.Range("M" & x).Value
                    tickerVolume = ws.Range("J" & x).Value
                End If
                
            Next x
            
            ' Print the values to the new table
            
            ws.Range("Q2").Value = tickerIncrease
            ws.Range("R2").Value = greatestPercentIncrease
            ws.Range("R2") = Format(greatestPercentIncrease, "0.00%")
            
            ws.Range("Q3").Value = tickerDecrease
            ws.Range("R3").Value = greatestPercentDecrease
            ws.Range("R3") = Format(greatestPercentDecrease, "0.00%")
            
            ws.Range("Q4").Value = tickerVolume
            ws.Range("R4").Value = greatestStockVolume
            
        ' Auto fit the worksheet
            ws.Range("A:R").Columns.AutoFit
        
    Next ws

End Sub


