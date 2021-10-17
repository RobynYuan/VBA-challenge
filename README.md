# VBA-challenge
Sub alphabetical_testing()
# for looping through the sheet
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

For Each ws In ThisWorkbook.Worksheets
   ws.Activate


        Dim Ticker As String
          Dim TotalStockvolumn As Double
          LastRow = Cells(Rows.Count, 1).End(xlUp).Row
         TotalStockvolumn = 0
          Dim summaryTableRow As Integer
            summaryTableRow = 2
           Dim Rowcounter As Double
           Rowcounter = 1
 # For looping throught the table to intergrating the column witht the same ticker   
             For i = 2 To LastRow
          If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           ' create summary table
            Ticker = Cells(i, 1).Value
          ' Writing to summary table
            Range("I" & summaryTableRow).Value = Ticker
            Range("L" & summaryTableRow).Value = TotalStockvolumn
            Range("J" & summaryTableRow).Value = yearlychange
            Range("K" & summaryTableRow).Value = percentagechange
            Range("K" & summaryTableRow).NumberFormat = "0.01%"
  # for conditioning formatting column J based on the value >0 or <0
             If Range("J" & summaryTableRow).Value > 0 Then
                 Range("J" & summaryTableRow).Interior.ColorIndex = 4
          Else
                 Range("J" & summaryTableRow).Interior.ColorIndex = 3
          End If

            
    
           summaryTableRow = summaryTableRow + 1
             
            
            ' reset total
       TotalStockvolumn = 0
       
  # Creating the bonus table header
        Cells(2, 15).Value = " Greatest % Increase"
       Cells(3, 15).Value = "Greatest % Decrease"
       Cells(4, 15).Value = " Greastest Total Volumn"
       Cells(1, 16).Value = "Ticker"
       Cells(1, 17).Value = "Value"

 
       
       Min = Application.WorksheetFunction.Min(Columns("K"))
       Max = Application.WorksheetFunction.Max(Columns("K"))
       Totalmax = Application.WorksheetFunction.Max(Columns("L"))
   # For creating the bonus table
       For r = 2 To summaryTableRow
       If Cells(r, 11).Value = Min Then
       Cells(3, 16).Value = Cells(r, 9).Value
       Cells(3, 17).Value = Min
       ElseIf Cells(r, 11).Value = Max Then
       Cells(2, 16).Value = Cells(r, 9).Value
       Cells(2, 17).Value = Max
       ElseIf Cells(r, 12).Value = Totalmax Then
       Cells(4, 16).Value = Cells(r, 9).Value
       Cells(4, 17).Value = Totalmax

       
       End If
       Next r
       


        Else
  #  caculating yearly change and percentagechange and Totalstockvolumn
            Rowcounter = Rowcounter + 1
            TotalStockvolumn = TotalStockvolumn + Cells(i, 7).Value
            yearlychange = Cells(i + 1, 6).Value - Cells((i - Rowcounter + 2), 3).Value
            percentagechange = (Cells(i, 6).Value - Cells((i - Rowcounter + 2), 3).Value) / Cells((i - Rowcounter + 2), 3).Value
            

    End If


    Next i
# For creating the header of the summary table
    Cells(1, 9) = "ticker"
    Cells(1, 10) = " Yearly Change"
    Cells(1, 11) = "percentage Change"
    Cells(1, 12) = "Total Stock volumn"

    
    End If
    Next ws



starting_ws.Activate 'activate the worksheet that was originally active

End Sub
