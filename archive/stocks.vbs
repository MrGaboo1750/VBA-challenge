sub stocks()

        dim i as Double
        dim current_ticker as string
        dim output_row_index as integer
        dim open_value as Double
        dim close_value as Double
        dim yearly_change as Double
        dim last_row as double
        
        
        ' Initialize the current ticker to a blank string
        current_ticker = ""
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"


        'find the last row of the spreadsheet
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (last_row)
        
        ' Create an index counter for the output rows, which must be independent of the stock rows
        output_row_index = 2
        'Loop over all stock ticker rows in the spreadsheet, starting at 2 and ending at 70926
        for i = 2 to last_row
                ' If a new ticker symbol is encountered...
                If (Cells(i, 1) <> current_ticker) Then
                        'reset the total stock volume to 0
                        total_stock_vol = 0
                        current_ticker = Cells(i, 1).Value
                        Cells(output_row_index, 9).Value = current_ticker
                        open_value = Cells(i, 3).Value

                Elseif (right(Cells(i, 2).value, 4) = "1230") Then
                        close_value = Cells(i, 6)
                        ' Add the final stock volume on close
                        total_stock_vol = total_stock_vol + Cells(i, 7).Value
                        yearly_change = close_value - open_value
                        percent_change = yearly_change / open_value
                        'Format the percent change as a percent. Note that this returns a string.
                        percent_change = format(percent_change, "Percent")
                        'Output the data to the cells
                        Cells(output_row_index, 10).Value = yearly_change
                        Cells(output_row_index, 11).Value = percent_change
                        Cells(output_row_index, 12).Value = total_stock_vol
                        'Increment the output row index counter by 1 so that the next ticker will be written to the next output row
                        output_row_index = output_row_index + 1
                
                ' Otherwise if the row is neither a new ticker nor the closing date, just add to the total stock volume
                Else
                        total_stock_vol = total_stock_vol + Cells(i, 7).Value


                End if
        next i
                        


end sub