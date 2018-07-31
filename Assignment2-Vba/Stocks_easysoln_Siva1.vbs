
Sub stocktotals()

'VBA script to determine totals of each ticker and write to columns I(ticker), J(Totals).

'Define variables
Dim ws As Worksheet
Dim wsfirst As Worksheet

Dim curstockname, hldstockname, nxtstockname As String
Dim wsname  As String

Dim stocktotal, curstockval, nxtstockval As Long
Dim LastRow As Long
Dim resultrow, cntofrows As Long
Dim cntofws  As Integer

'initialize relevant variables
cntofws = 0
cntofrows = 0
resultrow = 2
stocktotal = 0
 
'loop for each worksheet
For Each ws In Worksheets
    wsname = ws.Name
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox ("worksheet " + wsname)
    'MsgBox ("Last row in the worksheet is " + Str(LastRow))
    
    resultrow = 2
    cntofws = cntofws + 1
    stocktotal = 0
      
    'Populate titles for I and J.
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Volume"
         
    'loop for each row in the worksheet
    For i = 2 To LastRow
        'keep count total rows
        cntofrows = cntofrows + 1
        'assign values of ticker and total stock trades from cell values to variables.
        curstockname = ws.Cells(i, 1).Value
        curstockval = ws.Cells(i, 7).Value
        nxtstockname = ws.Cells(i + 1, 1).Value
        nxtstockval = ws.Cells(i + 1, 7).Value
        
        
        If i = 2 Then
            hldstockname = ws.Range("A2").Value
        End If
                
        If curstockname = nxtstockname Then
           stocktotal = stocktotal + curstockval
           
        Else
            'MsgBox (hldstockname)
            
            ' Print the ticker name  in the Result Table.
            ws.Range("I" & resultrow).Value = hldstockname
            
            'Add value of the last cell in that ticker to the stocktotal
            stocktotal = stocktotal + curstockval

            ' Print the Total stock value into the Result Table.
            ws.Range("J" & resultrow).Value = stocktotal
            
            'reset relevant values for next ticker.
            stocktotal = 0
            hldstockname = nxtstockname
            resultrow = resultrow + 1
        
        End If

    Next i

Next ws

MsgBox ("Total # of worksheets " + Str(cntofws))


End Sub




