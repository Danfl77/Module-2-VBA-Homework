Attribute VB_Name = "Module1"
Sub StockAnalysis()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets


    'Lable output columns and tables
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Value"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
   
    'Determine size of dataset on spreadsheet tab
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
     
    'Perform calculations and populate the main output table
    OutputTableRow = 2                                                              'Tracks the size of the output table
    OpenValue = ws.Range("C2").Value                                                'Determines the open value of the first stock
    Volume = ws.Range("G2")                                                         'Captures volume of first stock on first day
       
    For RowCounter = 2 To LastRow                                                   'Tracks progress through data set
        If ws.Cells(RowCounter + 1, 1).Value <> ws.Cells(RowCounter, 1) Then        'Locates transition from one ticker to the next
            ws.Cells(OutputTableRow, 9).Value = ws.Cells(RowCounter, 1).Value       'Places ticker name in output table
            CloseValue = ws.Cells(RowCounter, 6).Value                              'Determine close value
            YearlyChange = CloseValue - OpenValue                                   'Calculates yearly change
            ws.Cells(OutputTableRow, 10).Value = YearlyChange                       'Places yearly change in output table
            PercentChange = YearlyChange / OpenValue                                'Calculates percentage change
            ws.Cells(OutputTableRow, 11).Value = PercentChange                      'Places percent change in output table
            ws.Cells(OutputTableRow, 12).Value = Volume                             'Places stock volume in output table
            Volume = ws.Cells(RowCounter + 1, 7).Value                              'Reset the volume counter to the first value of next stock
            OpenValue = ws.Cells(RowCounter + 1, 3).Value                           'Resets opening value to that of the next stock
            OutputTableRow = OutputTableRow + 1
           
          Else
            Volume = Volume + ws.Cells(RowCounter + 1, 7).Value                     'If not transitioning to new ticker, continue to sum volume
        End If
    Next RowCounter
   
   
    'Conditional formatting of main output table
    TickerNameCount = ws.Cells(Rows.Count, 9).End(xlUp).Row - 1                     'Determine count of unique ticker names
       
    For OutputTableRow = 2 To (TickerNameCount + 1)
        ws.Cells(OutputTableRow, 11).Value = FormatPercent(ws.Cells(OutputTableRow, 11))  'Format column K as a percentage
        If ws.Cells(OutputTableRow, 10).Value < 0 Then                              'Determines if yearly change is negative
            ws.Cells(OutputTableRow, 10).Interior.ColorIndex = 3                    'Fills cell red if value is negative
        Else
            ws.Cells(OutputTableRow, 10).Interior.ColorIndex = 4                    'Fills cell green if value is positive
       End If
       
    Next OutputTableRow

   
   
    'Calculate greatest % increase
    HighestPercentIncrease = 0                                                      'Stores current highest percent increase
    Dim HigestIncreaseTicker As String
   
   
    For OutputTableRow = 2 To (TickerNameCount + 1)                                 'For loop determines if a value is larger the current higest percent increase and greater than the next percent in the table
        If ws.Cells(OutputTableRow, 11).Value > HighestPercentIncrease And ws.Cells(OutputTableRow, 11).Value > ws.Cells(OutputTableRow + 1, 11).Value Then
            HighestPercentIncrease = ws.Cells(OutputTableRow, 11).Value
            HigestIncreaseTicker = ws.Cells(OutputTableRow, 9).Value
         End If
       
    Next OutputTableRow
    ws.Range("Q2").Value = HighestPercentIncrease                                   'Places highest percent value in output table
    ws.Range("Q2").Value = FormatPercent(ws.Range("Q2"))                            'Formats highest percent value as a percentage
    ws.Range("P2").Value = HigestIncreaseTicker                                     'Places ticker name in the output table
           
           
    'Calculate greatest % decrease
    GreatestPercentDecrease = 0                                                     'Stores current greatest percent decrease
    Dim GreatestDecreaseTicker As String
               
    For OutputTableRow = 2 To (TickerNameCount + 1)                                 'For loop determines if a decrease is larger the current greatest percent decrease and larger than the next percent in the table
        If ws.Cells(OutputTableRow, 11).Value < GreatestPercentDecrease And ws.Cells(OutputTableRow, 11).Value < ws.Cells(OutputTableRow + 1, 11).Value Then
            GreatestPercentDecrease = ws.Cells(OutputTableRow, 11).Value
            GreatestDecreaseTicker = ws.Cells(OutputTableRow, 9).Value
         End If
     
     Next OutputTableRow
    ws.Range("Q3").Value = GreatestPercentDecrease                                   'Places greatest percent decrease value in output table
    ws.Range("Q3").Value = FormatPercent(ws.Range("Q3"))                             'Formats greatest percent decrease value as a percentage
    ws.Range("P3").Value = GreatestDecreaseTicker                                    'Places ticker name in output table
           
           
    'Calculate greatest total volume
    GreatestTotalVolume = 0                                                          'Stores current higest volume
    Dim TotalVolumeSticker As String                                                 'Stores current higest volume sticker name
   
    For OutputTableRow = 2 To (TickerNameCount + 1)                                  'Determines if value is higher than volume in the cell below and higher than current higest volume
        If ws.Cells(OutputTableRow, 12).Value > ws.Cells(OutputTableRow + 1, 12).Value And ws.Cells(OutputTableRow, 12).Value > GreatestTotalVolume Then
            GreatestTotalVolume = ws.Cells(OutputTableRow, 12).Value
            TotalVolumeSticker = ws.Cells(OutputTableRow, 9).Value
           
        End If
    Next OutputTableRow
    ws.Range("Q4").Value = GreatestTotalVolume
    ws.Range("P4").Value = TotalVolumeSticker
   
Next ws

End Sub


