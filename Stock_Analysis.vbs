
'Analyse generated stock market data

Sub Stock_Analysis()

Dim arr1 As Variant
Dim arr2 As Variant
Dim ws As Worksheet

'Loops through all the stocks for each quarter 
For Each ws In Worksheets

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Column header creation

arr1 = Array("Ticker", "Quaterly Change", "Percent Change", "Total Stock Volume")
ws.Range("I1:L1").Value = arr1

arr2 = Array("Ticker", "Value")
ws.Range("P1:Q1").Value = arr2

ticker = ws.Cells(2, 1).Value
qtrly_chng = 0
tot_vol = 0
open_val = ws.Cells(2, 3).Value
y = 2

ws.Range("k:k").NumberFormat = "0.00%"

For x = 2 To LastRow
        
    'check if ticker vlaue is changed
    If ws.Cells(x, 1).Value <> ticker Then
        
        qtrly_chng = ws.Cells(x - 1, 6).Value - open_val
        percent_chng = (qtrly_chng / open_val)
        
        ws.Cells(y, 9).Value = ticker
        ws.Cells(y, 10).Value = qtrly_chng
        ws.Cells(y, 11).Value = percent_chng
        ws.Cells(y, 12).Value = tot_vol

        'Conditional formatting that will highlight positive change in green and negative change in red

        If qtrly_chng > 0 Then
        
        ' Set the Cell Colours to Green
         ws.Cells(y, 10).Interior.ColorIndex = 4
  
         ElseIf qtrly_chng < 0 Then
         ' Set the Cell Colours to Red
          ws.Cells(y, 10).Interior.ColorIndex = 3
                
         End If
        
        open_val = ws.Cells(x, 3).Value
        tot_vol = ws.Cells(x, 7).Value
        ticker = ws.Cells(x, 1).Value
        y = y + 1
        
        'If ticker vlaue is not chnaged then keep adding total volume
    Else
        
        tot_vol = ws.Cells(x, 7).Value + tot_vol
       
    End If
    
Next x

'Return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

Max_val = Application.WorksheetFunction.Max(ws.Range("K:K"))
Min_val = Application.WorksheetFunction.Min(ws.Range("K:K"))
Max_vol = Application.WorksheetFunction.Max(ws.Range("L:L"))

lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

For Z = 2 To lastrow2

If ws.Cells(Z, 11).Value = Max_val Then

ws.Cells(2, 16).Value = ws.Cells(Z, 9)
ws.Cells(2, 17).Value = Format(Max_val, "Percent")

ElseIf ws.Cells(Z, 11).Value = Min_val Then

ws.Cells(3, 16).Value = ws.Cells(Z, 9)
ws.Cells(3, 17).Value = Format(Min_val, "Percent")

End If

If ws.Cells(Z, 12) = Max_vol Then

ws.Cells(4, 16).Value = ws.Cells(Z, 9)

End If

ws.Cells(4, 17).Value = Max_vol

Next Z

Next ws

End Sub
