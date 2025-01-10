Sub testing_VBA()
Dim i, j As Integer
Dim ws As Worksheet


    For Each ws In Worksheets   'For loop for going through every tap'
    
    Dim lastRowA As Long
    Dim opening_price As Double
    Dim closing_price As Double
    Dim Ticket As String
    Dim percentage_change As Double
    Dim total_stock As Double
    Dim lastRowH As Long
    Dim nextRowH As Long
    Dim lastRowI As Long
    Dim nextRowI As Long
    total_stock = 0
    
    
    lastRowA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
   
    'Creating Headers in the file'
    ws.Range("h1").Value = "Ticker"
    ws.Range("i1").Value = "Quarterly change"
    ws.Range("j1").Value = "Percentage Change"
    ws.Range("k1").Value = "Total stock volume"
    ws.Range("n2").Value = "Greatest % increase"
    ws.Range("n3").Value = "Greatest % decrease"
    ws.Range("n4").Value = "Greatest total volume"
    ws.Range("o1").Value = "Ticker"
    ws.Range("p1").Value = "Value"
    
    'Coping and pasting the first Ticker'
    ws.Select
    Range("A2").Select
    Selection.Copy
    Range("H2").Select
    ActiveSheet.Paste
    
    'For loop for going through the first column, skiping the headers'
    For i = 2 To lastRowA
        lastRowH = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row 'Difining the last available cell to paste the Ticket'
        nextRowH = lastRowH + 1
        
        If i = 2 Then  'Defining 1st opening price on the first ticker'
            opening_price = ws.Cells(i, 3).Value
        End If
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then  'If criteria for compare <ticker> column'
            ws.Select
            ws.Cells(i + 1, 1).Select
            Selection.Copy
            ws.Cells(nextRowH, "H").Select
            ActiveSheet.Paste
            closing_price = ws.Cells(i, 6).Value
            ws.Cells(nextRowH - 1, "I").Value = closing_price - opening_price  'Quarterly change calculation'
            ws.Cells(nextRowH - 1, "J").Value = (closing_price - opening_price) / opening_price  'Percentage Change calculation '
            opening_price = ws.Cells(i + 1, 3)
            'Applying color conditional formatting to column Quarterly change '
            If ws.Cells(nextRowH - 1, "I").Value > 0 Then
                ws.Cells(nextRowH - 1, "I").Interior.Color = vbGreen
            Else
               If ws.Cells(nextRowH - 1, "I").Value < 0 Then
                ws.Cells(nextRowH - 1, "I").Interior.Color = vbRed
                Else
                ws.Cells(nextRowH - 1, "I").Interior.Color = vbWhite
                End If
            End If
            'Loop for calculation Total stock volume'
            total_stock = 0
                For k = 2 To i
                    If ws.Cells(k, 1).Value = ws.Cells(i, 1).Value Then
                        total_stock = total_stock + ws.Cells(k, 7).Value
                    End If
                Next k
                ws.Cells(nextRowH - 1, "K").Value = total_stock
                
        End If

    

  Next i
    'For loop for calculation of stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"'
  
     For j = 2 To lastRowH
            If Cells(j, 10).Value = Application.WorksheetFunction.Max(ws.Range("J2:J" & lastRowH)) Then
                Cells(2, 15).Value = Cells(j, 8).Value
                Cells(2, 16).Value = Cells(j, 10).Value
                
            ElseIf Cells(j, 10).Value = Application.WorksheetFunction.Min(ws.Range("J2:J" & lastRowH)) Then
                Cells(3, 15).Value = Cells(j, 8).Value
                Cells(3, 16).Value = Cells(j, 10).Value
                
            ElseIf Cells(j, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastRowH)) Then
                Cells(4, 15).Value = Cells(j, 8).Value
                Cells(4, 16).Value = Cells(j, 11).Value
            End If
        Next j
        
        'Adding format to column I,J and cell with Greatest % increase", "Greatest % decrease"'
        ws.Range("I2:I" & lastRowH).NumberFormat = "0.00"
        ws.Range("J2:J" & lastRowH).NumberFormat = "0.00%"
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).NumberFormat = "0.00%"



    
    Next ws



End Sub

