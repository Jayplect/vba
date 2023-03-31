# Challenge2-VBA
### Description
The aim of this challenge was to create a script using VBA that loops through all multi-year stocks (each year in a separate worksheet)... 

<img width="358" alt="image" src="https://user-images.githubusercontent.com/107348074/228932543-56447522-f2c9-4c75-bb75-fd745e24ca91.png">

The outputs for this challenge includes a summary table showing: 
 ##### i) each unique ticker id across rows using the script below:
    For i = 2 To lRow
    'If next cell is not equal to preceding then...
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'retrieve the preceding cell value
         closingPrice = ws.Cells(i, 3)
        
        'store the preceding cell text (i.e., ticker label) in desired cell
         ws.Range("I" & selectrow).Value = ws.Cells(i, 1).Value
         
##### ii) other corresponding informations such as the yearly change for the opening price, percentage change from the opening price and the total stock volume for each ticker

A functionality was added to retrieve the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume":


            'Calculate the stock with the greatest increase, greatest decrease and greatest total volume
              For i = 2 To lRow_summary

              'Set variables to hold values
                MaxPrice = WorksheetFunction.Max(ws.Range("K2" & ":" & "K" & lRow_summary))
                MinPrice = WorksheetFunction.Min(ws.Range("K2" & ":" & "K" & lRow_summary))
                MaxVolume = WorksheetFunction.Max(ws.Range("L2" & ":" & "L" & lRow_summary))

              'Return the adjacent row to the specified cell
              '--------------------------------------------
                'Greatest % Increase
                  If ws.Cells(i, 11) = MaxPrice Then
                     ws.Range("P2").Value = MaxPrice
                     ws.Range("O2").Value = ws.Cells(i, 9)
                  End If
               '--------------------------------------------
                'Greatest % Decrease
                   If ws.Cells(i, 11) = MinPrice Then
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

#### A reset function was also included to easily clear the summary table in one click using the script:
        
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

#### A snapshot of the final result showing the summary table is presented below:
All stocks with a yearly change less than 0 are formatted to red while those greater than 0 are formatted green.

<img width="354" alt="image" src="https://user-images.githubusercontent.com/107348074/228947608-ea2410ff-cf75-486c-8ff8-cc93132af784.png">

### Reference
Data for this dataset was generated by edX Boot Camps LLC, and is intended for educational purposes only.
