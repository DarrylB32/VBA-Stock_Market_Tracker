
'STOCK MARKET HOMEWORK PLUS(Final)
Sub StockMarketPlus()
Dim ticker As Double 'tracks the ticker position
Dim open_amt As Double 'opening amount for the year
Dim close_amt As Double 'closing amount for the year
Dim volume As Double 'volume for a company
Dim Match As Boolean 'tells you if there is a match in the "ticker" column or not
Dim WS_Count As Integer 'identifies the number of worksheets
Dim w As Integer 'identifies the CURRENT worksheet used in the loop
Dim i As Double 'used to identify the row in the original data set and used for the greater percentages at the end
Dim j As Double ''used to identify the "ticker" row data set and used for the greater percentages at the end
WS_Count = ActiveWorkbook.Worksheets.Count ' Set WS_Count equal to the number of worksheets in the active workbook
For w = 1 To WS_Count 'Begin the loop for worksheets

Cells(1, 9).Value = "Ticker"        'ticker header
Cells(1, 10).Value = "Yearly Change"        'yearly change header
Cells(1, 11).Value = "Percent Change"       'percent change header
Cells(1, 12).Value = "Total Stock Volume"       'total volume header
Cells(1, 17).Value = "Ticker"       'second ticker header
Cells(1, 18).Value = "Value"        'percentage value header
Cells(2, 16).Value = "Greatest % Increaase"     'greatest increase row designation
Cells(3, 16).Value = "Greatest % Decreaase"     'greatest decrease row designation
Cells(4, 16).Value = "Greatest Total Volume"    'greatest volume row designation


open_amt = 0        'initial opening amount
close_amount = 0    'initial close amount
ticker = 0      'initial ticker position
volume = 0      'initial volume amount
For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row      'idientifies number of rows on current worksheet
'For i = 704932 To 705714
    Match = False       'match starts off as false
    For j = 2 To ticker + 2     'ticker row starts at 2
        If Cells(i, 1).Value = Cells(j, 9).Value Then       'if current company name is in ticker column then...
            volume = volume + Cells(i, 7).Value 'Add current volume to existing volume
            Match = True 'match has been found
                        
        End If
    Next j 'check next value in ticker column

    If Match = False Then       'if there is no match in ticker column
            If ticker <> 0 Then     'and ticker value is not equal to zero
                close_amt = Cells(i - 1, 6).Value ' use the ABOVE CELL of current "i" position to get the close amount
                    
                    If open_amt = 0 Then        'if open amount is equal to zero then...
                        open_amt = 1 'open amount will be made equal to 1 so we dont divide by a zero later on
                        close_amt = 1   'close amount will be equal to one
                        Cells(ticker + 1, 8).Value = "Open Amount was Zero. Please revise." 'place this string in cell next to ticker value to help identify this as an open amount of zero
                        Cells(ticker + 1, 8).Interior.ColorIndex = 6        'turn this cell yellow to help with identifying the cell with a close amount of zero
                        End If 'end what to do when the open amount is equal to zero
                        
                Cells(ticker + 1, 10).Value = close_amt - open_amt 'Update yearly change
                Cells(ticker + 1, 11).Value = (Cells(ticker + 1, 10).Value / open_amt) 'Update % changed
                Cells(ticker + 1, 11).NumberFormat = "0.00%" 'format cell as percentage
                Cells(ticker + 1, 12) = volume 'update volume
                    If Cells(ticker + 1, 10).Value < 0 Then 'if yearly change is negative then....
                        Cells(ticker + 1, 10).Interior.ColorIndex = 3       'make cell red
                            Else: Cells(ticker + 1, 10).Interior.ColorIndex = 4     'otherwise, make cell green since its positive
                    End If 'end coloring of cells
            End If      'end what hapens if the ticker value is not equal to zero
        Cells(ticker + 2, 9) = Cells(i, 1).Value 'if a MATCH IS NOT PRESENT add new company to ticker
        volume = Cells(i, 7).Value 'if a MATCH IS NOT PRESENT establish volume
        open_amt = Cells(i, 3).Value 'if a MATCH IS NOT PRESENT establish opening amount
        ticker = ticker + 1 'add 1 to the ticker to account for newly established values
    End If 'end what to do if a MATCH IS NOT present
Next i
close_amt = Cells(i - 1, 6).Value ' use the ABOVE CELL of current "i" position to get the close amount
Cells(ticker + 1, 10).Value = close_amt - open_amt 'Update yearly change
Cells(ticker + 1, 11).Value = (Cells(ticker + 1, 10).Value / open_amt) 'Update % changed
Cells(ticker + 1, 11).NumberFormat = "0.00%" 'format cell as percentage
Cells(ticker + 1, 12) = volume 'update volume
        If Cells(ticker + 1, 10).Value < 0 Then 'if yearly change is negative then....
            Cells(ticker + 1, 10).Interior.ColorIndex = 3       'make cell red
                Else: Cells(ticker + 1, 10).Interior.ColorIndex = 4     'otherwise, make cell green since its positive
        End If 'end coloring of cells

Dim grt_percent_inc As Double   'greatest percent increase variable
Dim grt_percent_dec As Double   'greatest percent decrease variable
Dim grt_total_vol As Double     'greatest total volume variable
grt_percent_inc = 0     'initial greatest percent increase amount
grt_percent_dec = 0     'initial greatest percent decrease amount
grt_total_vol = 0       'initial total volume amount
Dim tik_per_inc As String       'company for greatest increase variable
Dim tik_per_dec As String       'company for greatest decrease variable
Dim tik_tot_vol As String       'company for greatest total volume variable

For i = 2 To Cells(Rows.Count, 9).End(xlUp).Row     'number of rows in TICKER column for loop

If Cells(i, 11).Value > grt_percent_inc Then        'if percentage change of current cell is greater than saved value then...
grt_percent_inc = Cells(i, 11).Value        'replace old value with new value
tik_per_inc = Cells(i, 9).Value         'replace old company with new company name
End If

If Cells(i, 11).Value < grt_percent_dec Then        'if percentage change of current cell is less than than saved value then...
grt_percent_dec = Cells(i, 11).Value        'replace old value with new value
tik_per_dec = Cells(i, 9).Value     'replace old company with new company name
End If

If Cells(i, 12).Value > grt_total_vol Then      'if total volume of current cell is greater than saved value then...
grt_total_vol = Cells(i, 12).Value      'replace old value with new value
tik_tot_vol = Cells(i, 9).Value     'replace old company with new company name
End If

Next i

Cells(2, 18).Value = grt_percent_inc    'place greatest percent increase value in location
Cells(3, 18).Value = grt_percent_dec    'place greatest percent decrease value in location
Cells(4, 18).Value = grt_total_vol  'place greatest volume value in location
Cells(2, 17).Value = tik_per_inc    'place greatest percent increase value company name in location
Cells(3, 17).Value = tik_per_dec    'place greatest percent decrease value company name in location
Cells(4, 17).Value = tik_tot_vol    'place greatest volume value comany name in location


'MsgBox ActiveWorkbook.Worksheets(i).Name       'create message box for complated worksheet
Worksheets(ActiveSheet.Index + 1).Select        'go to next worksheet to display values
Next w      'go to next worksheet to extract data


End Sub

