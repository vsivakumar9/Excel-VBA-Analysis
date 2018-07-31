

Sub stocktotals()

'VBA script to determine totals of each ticker and write to columns I(ticker),J(Yearly change), K(percent yearly change) L(Totals).

'Define variables
Dim ws As Worksheet
Dim wsfirst As Worksheet

Dim curstockname, hldstockname, nxtstockname As String
Dim wsname  As String

Dim stocktotal, curstockval, nxtstockval As Long
Dim LastRow As Long
Dim resultrow, cntofrows As Long

Dim cntofws  As Integer

Dim curopenval, curclosevalue As Double
Dim tkryearchng, tkrclosevalue, tkropenvalue, tkrpercentchng As Double
Dim Isfirstrow_ticker As Boolean


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
    MsgBox ("worksheet " + wsname)
    'MsgBox ("Last row in the worksheet is " + Str(LastRow))
    
    cntofws = cntofws + 1
    stocktotal = 0
    resultrow = 2
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"
    'set interior color to Yellow-6
    ws.Range("I1,J1,K1,L1").Interior.ColorIndex = 6
    ws.Range("I1,J1,K1,L1").Font.Bold = True
         
             
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
            
                       
            'Print yearcly percent change for ticker in column K.
            ws.Range("K" & resultrow).Value = tkrpercentchng
            'set format to % in the cell.
            ws.Range("K" & resultrow).NumberFormat = "0.0000%"
            
            
            ' Print the Total stock value into the Result Table.
            ws.Range("L" & resultrow).Value = stocktotal
            
            
            'reset relevant values for next ticker.
            stocktotal = 0
            hldstockname = nxtstockname
            resultrow = resultrow + 1
            Isfirstrow_ticker = True
        
        End If

    Next i

'color the percent change columns using conditional formatting.


Next ws

MsgBox ("Total # of worksheets " + Str(cntofws))


End Sub





