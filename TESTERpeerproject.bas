Attribute VB_Name = "Module2"
Sub BRIT_KEV_CAITE_JALOPY()
Attribute BRIT_KEV_CAITE_JALOPY.VB_ProcData.VB_Invoke_Func = " \n14"
'
' conversiontonumver Macro
'

'
End Sub
'create column headings
Sub Stock_market():
 Dim ws As Worksheet
 Dim ticker As String
 Dim i As Long
 Dim openprice As Double
 Dim Yearlychange As Double
 Dim PecentageChange As Double
 Dim Volume As LongLong
 Dim n As Long
'activating all worksheets
    For Each ws In Worksheets
'Find last row in worksheet
    Dim lastRow As Long
    lastRow = 5000 ' WS.Cells(Rows.Count, "A").End(xlUp).Row
'create Column headers
    ws.Range("i1").Value = "Ticker"
    ws.Range("j1").Value = "Yearly Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("l1").Value = " Total Stock Volume"
' add return the stock with Greatest & Increase, Greatest % Decrease and Greatest total volume
    ws.Range("p1").Value = "Ticker"
    ws.Range("q1").Value = "Value"
    ws.Range("o2").Value = "Greatest % Increase"
    ws.Range("o3").Value = "Greatest % Decrease"
    ws.Range("o4").Value = "Greatest Total Volume"
'Copy Ticker
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1).Value Then
                ws.Range("i" & ws.Range("i:i").Rows.Count).End(xlUp).Offset(1, 0) = ws.Range("a" & i)
                openprice = ws.Range("a" & i).Offset(0, 2)
'Yearly Change
            ElseIf ws.Range("a" & i) <> ws.Range("a" & i + 1) Then
                Yearlychange = ws.Range("a" & i).Offset(0, 5) - openprice
                ws.Range("j" & ws.Range("j:j").Rows.Count).End(xlUp).Offset(1, 0) = Yearlychange
                ws.Range("k" & ws.Range("k:k").Rows.Count).End(xlUp).Offset(1, 0) = Yearlychange / openprice
        End If
         ws.Range("k:k").NumberFormat = ("0.00%")
 'Total Stock Volume
       Volume = 0
        n = 2
        If ws.Range("a" & i) = ws.Range("i" & n) Then
            Volume = ws.Range("a" & i).Offset(0, 6) + Volume
            ElseIf ws.Range("a" & i) <> ws.Range("a" & i + 1) Then
            ws.Range("l" & ws.Range("l:l").Rows.Count).End(xlUp).Offset(1, 0) = Volume
            n = n + 1
            Else: n = n + ws.Cells(i, 3).Value
        End If
   Next i
Next ws
End Sub

