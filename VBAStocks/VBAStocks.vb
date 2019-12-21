Sub vbaStocks()

For Each ws In Worksheets

    'Print column headers for summary tables
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Yearly % Change"
    ws.Range("L1").Value = "Total Volume"

    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Amount"

    'Define data set indexes
        firstRowIndex = 2
        tickerColIndex = 1
        openColIndex = 3
        closeColIndex = 6
        volColIndex = 7
        lastRowIndex = ws.Cells(Rows.Count, "A").End(xlUp).Row

    'Define primary summary table indexes
        sum_firstRowIndex = 1 'start @ header then add 1 each iteration
        sum_tickerColIndex = 9
        sum_yrlyChangeColIndex = 10
        sum_yrlyPercentChangeColIndex = 11
        sum_totalVolColIndex = 12

    'Initial value for total volume summation
        totalVolume = 0

    'Create primary summary table
        For i = firstRowIndex To lastRowIndex
                
                previousTicker = ws.Cells(i - 1, 1).Value
                currentTicker = ws.Cells(i, 1).Value
                nextTicker = ws.Cells(i + 1, 1).Value

            If currentTicker <> previousTicker Then 'Condtion for first instance of a ticker symbol
                
                        'Assign yearly open price and begin summing daily volume
                        yrlyOpen = ws.Cells(i, openColIndex)
                        totalVolume = totalVolume + ws.Cells(i, volColIndex)
            
            ElseIf previousTicker = nextTicker Then 'Condtion for all instances of ticker symbol, except for the first and last instance
                
                        'Continue summing daily volume
                        totalVolume = totalVolume + ws.Cells(i, volColIndex)
                
            Else 'Condtion for last instance of a ticker symbol
                
                        'Continue summing daily volume, assign yearly close price and assign current ticker symbol's row index for the primary summary table
                        totalVolume = totalVolume + ws.Cells(i, volColIndex)
                        yrlyClose = ws.Cells(i, closeColIndex)
                        sum_firstRowIndex = sum_firstRowIndex + 1

                        'Calculate yearly change in price
                        yrlyChange = yrlyClose - yrlyOpen


                        'Calculate yearly % change in price using if statement to prevent #DIV/0 error
                        If yrlyOpen = 0 Then

                            yrlyPercentChange = 0

                        Else

                            yrlyPercentChange = yrlyChange / yrlyOpen

                        End If


                        'Print current ticker symbol's info into primary summary table
                        ws.Cells(sum_firstRowIndex, sum_tickerColIndex) = currentTicker
                        ws.Cells(sum_firstRowIndex, sum_yrlyChangeColIndex) = yrlyChange
                        ws.Cells(sum_firstRowIndex, sum_yrlyPercentChangeColIndex) = yrlyPercentChange
                        ws.Cells(sum_firstRowIndex, sum_totalVolColIndex) = totalVolume

                        'Reset total volume summation before moving to the next ticker symbol
                        totalVolume = 0
                

            End If

        Next i
    
    'Add conditional formating to yrly change column
        sum_lastRowIndex = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        For j = 2 To sum_lastRowIndex

            If ws.Cells(j, sum_yrlyChangeColIndex) > 0 Then

                ws.Cells(j, sum_yrlyChangeColIndex).Interior.Color = RGB(0, 255, 0)

            Else 

                ws.Cells(j, sum_yrlyChangeColIndex).Interior.Color = RGB(255, 0, 0)

            End If

        Next j

    'Create secondary summary table

        'Define ranges from primary summary table (to be used in secondary table)
        sum_tickerCol = ws.Range("I:I")
        sum_yrlyPercentChangeCol = ws.Range("K:K")
        sum_totalVolCol = ws.Range("L:L")
    
        'Find max/min values
        greatestPerInc = Application.WorksheetFunction.Max(sum_yrlyPercentChangeCol)
        greatestPerDec = Application.WorksheetFunction.Min(sum_yrlyPercentChangeCol)
        greatedTotalVol = Application.WorksheetFunction.Max(sum_totalVolCol)

        'Find corresponding ticker symbol using index/match function
        greatestPerIncTicker = WorksheetFunction.Index(sum_tickerCol, WorksheetFunction.Match(greatestPerInc, sum_yrlyPercentChangeCol, 0))
        greatestPerDecTicker = WorksheetFunction.Index(sum_tickerCol, WorksheetFunction.Match(greatestPerDec, sum_yrlyPercentChangeCol, 0))
        greatedTotalVolTicker = WorksheetFunction.Index(sum_tickerCol, WorksheetFunction.Match(greatedTotalVol, sum_totalVolCol, 0))

        'Print outputs into secondary summary table
        ws.Range("O2").Value = greatestPerIncTicker
        ws.Range("P2").Value = greatestPerInc

        ws.Range("O3").Value = greatestPerDecTicker
        ws.Range("P3").Value = greatestPerDec

        ws.Range("O4").Value = greatedTotalVolTicker
        ws.Range("P4").Value = greatedTotalVol

Next ws

End Sub



