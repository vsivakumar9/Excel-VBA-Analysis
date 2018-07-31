
Sub stocktotals()

'VBA script to determine totals of each ticker and write to columns I(ticker),J(Yearly change), K(percent yearly change) L(Totals).
'Also identify stock ticker with max increase, maximum decrease and max total volume.

'Define variables
Dim ws As Worksheet
Dim wsfirst As Worksheet

Dim curstockname, hldstockname, nxtstockname As String
Dim wsname  As String

Dim curstockval, nxtstockval As Long
Dim LastRow As Long
Dim resultrow, cntofrows, result_greatrow As Long
Dim stocktotal As Double


Dim cntofws  As Integer

Dim curopenval, curclosevalue As Double
Dim tkryearchng, tkrclosevalue, tkropenvalue, tkrpercentchng As Double
Dim Isfirstrow_ticker As Boolean

'* Variables to store max increase, max decrease  and max volume.
Dim Maxincrease_ticker, Maxdecrease_ticker, Maxvolume_ticker As String
Dim Maxincrease_value, Maxdecrease_value
'Dim Maxvolume_value As Long
Dim Maxvolume_value As Double

'initialize relevant variables
cntofws = 0
cntofrows = 0

stocktotal = 0
tkryearchng = 0
tkrpercentchng = 0

Isfirstrow_ticker = True

'loop for each worksheet
For Each ws In Worksheets
    wsname = ws.Name
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox ("Starting worksheet " + wsname)
    'MsgBox ("Last row in the worksheet is " + Str(LastRow))
    
    cntofws = cntofws + 1
    stocktotal = 0
    resultrow = 2
    resultgreatrow = 2
    
    'Init of variables to store max increase, max decrease  and max volume.
    Maxincrease_value = 0
    Maxdecrease_value = 0
    Maxvolume_value = 0
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest Total volume"
    
    'set interior color to Yellow-6
    ws.Range("I1,J1,K1,L1,P1,Q1").Interior.ColorIndex = 6
    ws.Range("I1,J1,K1,L1,P1,Q1").Font.Bold = True
         
             
    'loop for each row in the worksheet
    For i = 2 To LastRow
        'keep count total rows
        cntofrows = cntofrows + 1
        'assign values of ticker and total stock trades from cell values to variables.
        curstockname = ws.Cells(i, 1).Value
        curstockval = ws.Cells(i, 7).Value
        curopenval = ws.Cells(i, 3).Value
        curcloseval = ws.Cells(i, 6).Value
        
        nxtstockname = ws.Cells(i + 1, 1).Value
        nxtstockval = ws.Cells(i + 1, 7).Value
        
        'capture the open value on first day.
        If Isfirstrow_ticker = True Then
            tkropenvalue = curopenval
            
            'set value to false for other rows.
            Isfirstrow_ticker = False
        End If
        
        
        If i = 2 Then
            hldstockname = ws.Range("A2").Value
        End If
                
        If curstockname = nxtstockname Then
           stocktotal = stocktotal + curstockval
           
        Else
            'MsgBox (hldstockname)
            '*** This is the break condition when current row ticker not = next row ticker.
            'Add value of the last cell in that ticker to the stocktotal
            stocktotal = stocktotal + curstockval
            
            tkrclosevalue = curcloseval
            
            tkryearchng = tkrclosevalue - tkropenvalue
            
            'calculate percent change. set % value to )  when denominator zero to avoid divide by  0 error.
            If tkropenvalue <> 0 Then
               tkrpercentchng = tkryearchng / tkropenvalue
            Else
                tkrpercentchng = 0
            End If
            
            ' Print the ticker name  in the Result Table.
            ws.Range("I" & resultrow).Value = hldstockname
            
            'Print yearcly change for ticker in column J.
            ws.Range("J" & resultrow).Value = tkryearchng
            If tkryearchng >= 0 Then
                'set background color to green if positive change.
                ws.Range("J" & resultrow).Interior.ColorIndex = 4
                
            Else
                'set background color to red if negative change.
                ws.Range("J" & resultrow).Interior.ColorIndex = 3
            End If
            
            'if percent increase of cur row  > saved percent increase, update saved
            'value of max  increase.
            
            'if  percent decrease  of cur row  < saved percent decrease, update saved
            'value of max  decrease.
            If tkrpercentchng > 0 Then
                If tkrpercentchng > Maxincrease_value Then
                    Maxincrease_value = tkrpercentchng
                    Maxincrease_ticker = hldstockname
                End If
                
            ElseIf tkrpercentchng < 0 Then
                If tkrpercentchng < Maxdecrease_value Then
                    Maxdecrease_value = tkrpercentchng
                    Maxdecrease_ticker = hldstockname
                End If
            
            End If
                        
            'Print yearcly percent change for ticker in column K.
            ws.Range("K" & resultrow).Value = tkrpercentchng
            'set format to % in the cell.
            ws.Range("K" & resultrow).NumberFormat = "0.0000%"
                       
                        
            ' Print the Total stock value into the Result Table.
            ws.Range("L" & resultrow).Value = stocktotal
            
            'if total volume of cur row > saved total volume, update saved total vol.
            If stocktotal > Maxvolume_value Then
            '    MsgBox (Str(stocktotal))
            '    MsgBox (Str(Maxvolume_value))
            '   Maxvolume_value_d = double(stocktotal)
                Maxvolume_value = stocktotal
                Maxvolume_ticker = hldstockname
            
            End If
            
            'reset relevant values for next ticker.
            stocktotal = 0
            hldstockname = nxtstockname
            resultrow = resultrow + 1
            Isfirstrow_ticker = True
        
        End If

    Next i

'Populate max increase, max decrease and max total vaolume columns.

ws.Range("P2").Value = Maxincrease_ticker
ws.Range("Q2").Value = Maxincrease_value
ws.Range("Q2").NumberFormat = "0.0000%"

ws.Range("P3").Value = Maxdecrease_ticker
ws.Range("Q3").Value = Maxdecrease_value
ws.Range("Q3").NumberFormat = "0.0000%"

ws.Range("P4").Value = Maxvolume_ticker
ws.Range("Q4").Value = Maxvolume_value


Next ws

MsgBox ("Total # of worksheets " + Str(cntofws))


End Sub





