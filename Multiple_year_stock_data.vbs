Attribute VB_Name = "Module1"
Sub loopscript():
'-----------------------------------------------------------------------------------
'LOOP THROUGH ALL SHEETS
'-----------------------------------------------------------------------------------
 
 For Each ws In Worksheets

    'Determine the length of the first column
     Dim lRow As Double
     lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
     'keep track of the location for each ticker in the summary table
     Dim selectrow As Double
     selectrow = 2
     
     'initialize a variable to hold iteration when the Else condition below is met
     Dim iteration As Double
     iteration = 0
     
     'Set an initial variable for holding the total stock volume for each unique ticker
     Dim totalVolume As Double
     totalVolume = 0
     
 '-----------------------------------------------------------------------------------
 'INCLUDE A HEADERS FOR SUMMARY TABLE
 '-----------------------------------------------------------------------------------
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "YearlyChange"
    ws.Range("K1").Value = "PercentChange"
    ws.Range("L1").Value = "TotalStockVolume"
    ws.Range("N1").Value = "Parameters"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
     
 '-----------------------------------------------------------------------------------
 'EXECUTE THE LOOP FOR EACH WORSHEET
 '-----------------------------------------------------------------------------------
 
  For i = 2 To lRow
    'If next cell is not equal to preceding then...
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'retrieve the preceding cell value
         closingPrice = ws.Cells(i, 3)
        
        'store the preceding cell text (i.e., ticker label) in desired cell
         ws.Range("I" & selectrow).Value = ws.Cells(i, 1).Value
        
        'Calculate the yearly change
         yearlyChange = closingPrice - openingPrice
         
         'Calculate the percentage change
         percentChange = ((closingPrice - openingPrice) / openingPrice) * 100
         
         ' Add to the total stock volume
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        
        'Store the yearly change,  percentage change and total stock volumes values in desired cells respectively:
         ws.Range("J" & selectrow).Value = yearlyChange
         ws.Range("K" & selectrow).Value = percentChange
         ws.Range("L" & selectrow).Value = totalVolume
          
         'Increase the selected row by 1
         selectrow = selectrow + 1
         
         'Reset the total volume
         totalVolume = 0
     
   
  'If next row cell is equal to preceding row cell then...
  Else
      
       'Grep the opening price and store it in variable: openingPrice
       openingPrice = ws.Cells(i + iteration, 3).Value
       
       'upon every iteration under the Else condition, subtract 1 from the iteration
       iteration = iteration - 1
       
        'sum the stock volume
       totalVolume = totalVolume + ws.Cells(i, 7).Value
    
   
   End If
   
   
  Next i
     
 '-----------------------------------------------------------------------------------
 'FORMAT YEARLY CHANGE CELLS
 '-----------------------------------------------------------------------------------
  'Determine the length of the summary table
 Dim lRow_summary As Double
 lRow_summary = ws.Cells(Rows.Count, "J").End(xlUp).Row
     
  'Loop through the summary table
  For i = 2 To lRow_summary
    If ws.Range("J" & i).Value < 0 Then
    ws.Range("J" & i).Interior.ColorIndex = 3
    
  Else
    ws.Range("J" & i).Interior.ColorIndex = 4
  
   End If
  
  Next i
  
'-----------------------------------------------------------------------------------
'CALCULATED VALUES
'-----------------------------------------------------------------------------------
'Calculate the stock with the greatest increase, greatest decrease and greatest total volume
For i = 2 To lRow_summary

'Set variables to hold values
  MaxPrice = WorksheetFunction.Max(ws.Range("J2" & ":" & "J" & lRow_summary))
  MinPrice = WorksheetFunction.Min(ws.Range("J2" & ":" & "J" & lRow_summary))
  MaxVolume = WorksheetFunction.Max(ws.Range("L2" & ":" & "L" & lRow_summary))

'Return the adjacent row to the specified cell
'--------------------------------------------
  'Greatest % Increase
    If ws.Cells(i, 10) = MaxPrice Then
       ws.Range("P2").Value = MaxPrice
       ws.Range("O2").Value = ws.Cells(i, 9)
    End If
 '--------------------------------------------
  'Greatest % Decrease
     If ws.Cells(i, 10) = MinPrice Then
       ws.Range("P3").Value = MinPrice
       ws.Range("O3").Value = ws.Cells(i, 9)
    End If
 '--------------------------------------------
  'Greatest total volume
    If ws.Cells(i, 12) = MaxVolume Then
       ws.Range("P4").Value = MaxVolume
       ws.Range("O4").Value = ws.Cells(i, 9)
    End If
 '--------------------------------------------
'set the title parameters
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"

  Next i
 
 Next ws
 
 MsgBox ("Great job! Script ran without bugs!")
  
End Sub

'-----------------------------------------------------------------------------------
'RESET THE SUMMARY TABLE
'-----------------------------------------------------------------------------------

Sub reset():
'set the summary table to blank
For Each wsSummary In Worksheets
    wsSummary.Range("I:I").Value = ""
    wsSummary.Range("J:J").Value = ""
    wsSummary.Range("K:K").Value = ""
    wsSummary.Range("L:L").Value = ""
    wsSummary.Range("N1:P4").Value = ""

'Format yearly price summary color
    wsSummary.Range("J:J").Interior.ColorIndex = 0
    
 Next wsSummary
 
End Sub

