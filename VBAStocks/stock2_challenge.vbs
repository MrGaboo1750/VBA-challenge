
sub stocks2()

' Execute this code for every worksheet
        For each wksht in Worksheets

                ' Initialize the Ticker, Yearly Change, Percent Change, and Total Stock Volume output rows
                wksht.Range("I1").Value = "Ticker"
                wksht.Range("J1").Value = "Yearly Change"
                wksht.Range("K1").Value = "Percent Change"
                wksht.Range("L1").Value = "Total Stock Volume"

                'count number of rows in sheet
                numrows = wksht.Cells(Rows.Count, 1).End(xlUp).Row

                ' initialize a counter for the data output row
                output_row = 2

                ' loop over every row in the sheet, starting with the second row (the first row has headers)
                For row = 2 to numrows
                        ' compare the stock ticker of the current row with the ticker of the next row. If they are not the same, then you are finished counting stock volume for the current ticket.
                        If (wksht.Cells(row, 1).Value <> wksht.Cells(row + 1, 1).Value) Then

                                ' Add the final stock volume for this ticker to the running count
                                stock_vol = stock_vol + wksht.Cells(row, 7)

                                ' Report out the ticker symbol and the sum of the stock volume
                                wksht.Cells(output_row, 9).Value = wksht.Cells(row, 1).Value

                                ' When the ticker is changing on the next row, the closing price for the current row is that closing price for that year, since that is the last row with this ticker symbol. Set that closing price equal to closing_price 
                                closing_price = wksht.Cells(row, 6).Value

                                ' If the opening price is non-zero, proceed to calculate the yearly change and percent change
                                If (opening_price <> 0) Then
                                        ' Calculate the yearly change and percent change, then report them out to the output table. 
                                        yearly_change = closing_price - opening_price
                                        percent_change = yearly_change / opening_price
                                        percent_change = formatpercent(percent_change, 2)

                                        wksht.Cells(output_row, 10).Value = yearly_change

                                        ' If the yearly change is positive, highlight in green. If negative, highlight in red. If zero, do not highlight
                                        If (wksht.Cells(output_row, 10).Value > 0) Then
                                                wksht.Cells(output_row, 10).Interior.ColorIndex = 4

                                        Elseif (wksht.Cells(output_row, 10).Value < 0) Then
                                                wksht.Cells(output_row, 10).Interior.ColorIndex = 3
                                        End if

                                        wksht.Cells(output_row, 11).Value = percent_change
                                        wksht.Cells(output_row, 12).Value = stock_vol
                                ' If the opening price is zero, populate the output table with zeroes for the yearly change and percent change
                                Else
                                        wksht.Cells(output_row, 10).Value = 0
                                        wksht.Cells(output_row, 11).Value = 0
                                End if

                                ' Reset the stock_vol counter to 0 so that the next ticker gets a fresh start
                                stock_vol = 0


                                ' Increment the output row by 1 so that the next ticker is reported on the next row
                                output_row = output_row + 1

                                ' Set opening_price equal to the open price for the next row, as it is a new ticker. This opening price will be held in memory until it is overwritten during the next ticker mismatch
                                opening_price = wksht.Cells(row + 1, 3).Value

                        ' If the stock ticker for the next row is the same as the ticker for the current row, then we just add the stock volume for the current row to the running sum
                        Else
                                ' Add stock volume to the running count
                                stock_vol = stock_vol + wksht.Cells(row, 7)

                                'If this is the first row in the sheet, initialize the opening price as it does not yet exist. For all other rows this value will already be stored
                                If (row = 2) Then
                                        opening_price = wksht.Cells(row, 3).Value
                                End if

                        End if

                next row

                '15, 16, 17
                
                ' Initialize greatest % increase/decrease values, and greatest stock volume value
                greatest_percent_increase = 0
                greatest_percent_decrease = 0
                greatest_stock_vol = 0
                
                ' loop through the % change values in the output table and store the highest and lowest. The number of rows in the output table has already been stored as output_row, so use that variable
                
                For resultrow = 2 to output_row
                        ' If the total stock vol for this row is greater than the greatest stock volume calculate thus far, update the variable and store the ticker name as greatest_vol_ticker
                        If ( wksht.Cells(resultrow, 12).Value > greatest_stock_vol ) Then
                                greatest_stock_vol = wksht.Cells(resultrow, 12).Value
                                greatest_vol_ticker = wksht.Cells(resultrow, 9).Value
                        End If

                        ' If the percent change in this row is smaller than the greatest_percent_decrease, then update the variable and save the current ticker as biggest_loser
                        If ( (wksht.Cells(resultrow, 11).Value < greatest_percent_decrease) ) Then
                                greatest_percent_decrease = wksht.Cells(resultrow, 11).Value
                                biggest_loser = wksht.Cells(resultrow, 9).Value
                        ' If the percent change in this row is larger than the greatest_percent_increase, then update the variable and save the current ticker as biggest_gainer
                        Elseif ( (wksht.Cells(resultrow, 11).Value > greatest_percent_increase) ) Then
                                greatest_percent_increase = wksht.Cells(resultrow, 11).Value
                                biggest_gainer = wksht.Cells(resultrow, 9).Value
                        End If

                next resultrow

                ' output the greatest % increase/decrease and greatest total volume values

                ' initialize the table headings
                wksht.Range("P1").Value = "Ticker"
                wksht.Range("Q1").Value = "Value"
                wksht.Range("O2").Value = "Greatest % Increase"
                wksht.Range("O3").Value = "Greatest % Decrease"
                wksht.Range("O4").Value = "Greatest Total Volume"

                ' output the ticker symbols
                wksht.Range("P2").Value = biggest_gainer
                wksht.Range("P3").Value = biggest_loser
                wksht.Range("P4").Value = greatest_vol_ticker

                wksht.Range("Q2").Value = formatpercent(greatest_percent_increase, 2)
                wksht.Range("Q3").Value = formatpercent(greatest_percent_decrease, 2)
                wksht.Range("Q4").Value = greatest_stock_vol

                ' autofit all columns in worksheet
                wksht.Columns("A:Q").AutoFit
        
        next wksht

end sub