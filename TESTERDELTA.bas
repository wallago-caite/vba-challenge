Attribute VB_Name = "Module3"
Sub TESTERDELTA():
'activating all worksheets
    For Each ws In Worksheets

'Find last row in worksheet
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row


    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

' add return the stock with Greatest & Increase, Greatest % Decrease and Greatest total volume
    ws.Range("p1").Value = "Ticker"
    ws.Range("q1").Value = "Value"
    ws.Range("o2").Value = "Greatest % Increase"
    ws.Range("o3").Value = "Greatest % Decrease"
    ws.Range("o4").Value = "Greatest Total Volume"

    ' DECLARE PRICE Set an initial variable for holding the total per line and the yearly change
  Dim Total_Stock_Volume As LongLong
  Total_Stock_Volume = 0
  
  Dim openprice As Double
  Dim closedprice As Double
  Dim percentchange As Double
  Dim Yearlychange As Long
  
  'initial openprice we already know what the openprice is in the first cell
   
  openprice = ws.Cells(2, 3).Value
  closedprice = 0

  ' Keep track of the location for each criteria edit
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

   ' Loop through all tickers
  For i = 2 To lastRow

    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Set Ticker, filter for unique
            ticker = ws.Cells(i, 1).Value
            
                'set openprice
            openprice = ws.Cells(i, 3).Value
                
                'set closedprice
            closedprice = ws.Cells(i, 6).Value
            
             'closedprice-openprice= yearlychange
             Yearlychange = closedprice - openprice
        
             'find percentincrease
             percentchange = Yearlychange / openprice

                ' Add to the Total Stock Volume
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
             ' Print the ticker in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = ticker
                        
            'print yearlychange
            ws.Range("J" & Summary_Table_Row).Value = Yearlychange
       
            'print percentincrease
            ws.Range("K" & Summary_Table_Row).Value = percentchange
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
             
             ' Print the stock volume Amount to the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
             
            ' Add one to the summary table row
             Summary_Table_Row = Summary_Table_Row + 1
      
            ' Reset the Stock volume Total
             Total_Stock_Volume = 0
                ' If the cell immediately following a row is the same ticker..
            
            'Reset the openprice
            openprice = ws.Cells(i + 1, 3).Value
            
            'reset the closedprice
            closedprice = 0
      
        Else

         ' Add to the stock Total
             Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
    End If
    
Next i

'Find last row of second column
    Dim lastRow2 As Long
    lastRow2 = ws.Cells(Rows.Count, "I").End(xlUp).Row

For Z = 2 To lastRow2
    
    'change the colors in yearly and percentchange
        If ws.Cells(Z, 10).Value >= 0 Then
            ws.Cells(Z, 10).Interior.Color = vbGreen
            ws.Cells(Z, 11).Interior.Color = vbGreen
        ElseIf ws.Cells(Z, 10).Value < 0 Then
             ws.Cells(Z, 10).Interior.Color = vbRed
             ws.Cells(Z, 11).Interior.Color = vbRed
             
        End If

Next Z

'Report on maxes, mins
'greatest increase value and ticker then print
Dim greatestpercentincrease As Double
    greatestpercentincrease = Application.WorksheetFunction.max(ws.Range("K2:k5000"))
        ws.Range("Q2").Value = greatestpercentincrease
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
Dim tickerpercentin As String
    tickerpercentin = Application.WorksheetFunction.XLookup(ws.Range("q2"), ws.Range("K2:k5000"), ws.Range("I2:i5000"), 0, 0, 1)
        ws.Range("P2").Value = tickerpercentin

'greatest decrease value and ticker then print
Dim greatestpercentdecrease As Double
    greatestpercentdecrease = Application.WorksheetFunction.min(ws.Range("K2:k5000"))
        ws.Range("Q3").Value = greatestpercentdecrease

Dim tickerpercentdec As String
    tickerpercentdec = Application.WorksheetFunction.XLookup(ws.Range("Q3"), ws.Range("K2:k5000"), ws.Range("I2:I5000"), 0, 0, 1)
        ws.Range("P3").Value = tickerpercentdec

'greatest total value and ticker then print
Dim greatesttotal As LongLong
    greatesttotal = Application.WorksheetFunction.max(ws.Range("L2:L5000"))
        ws.Range("Q4").Value = greatesttotal

Dim tickergreattotal As String
    tickergreattotal = Application.WorksheetFunction.XLookup(ws.Range("Q4"), ws.Range("L2:L5000"), ws.Range("I2:I5000"), 0, 0, 1)
        ws.Range("P4").Value = tickergreattotal
Next ws

End Sub


