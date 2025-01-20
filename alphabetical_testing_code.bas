Attribute VB_Name = "Module1"
Sub testing_VBA()
Dim i, j As Integer
Dim ws As Worksheet

    For Each ws In Worksheets
    
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
   
    ws.Range("h1").Value = "Ticker"
    ws.Range("i1").Value = "Quarterly change"
    ws.Range("j1").Value = "Percentage Change"
    ws.Range("k1").Value = "Total stock volume"
    ws.Range("n2").Value = "Greatest % increase"
    ws.Range("n3").Value = "Greatest % decrease"
    ws.Range("n4").Value = "Greatest total volume"
    ws.Range("o1").Value = "Ticker"
    ws.Range("p1").Value = "Value"
    
    ws.Select
    Range("A2").Select
    Selection.Copy
    Range("H2").Select
    ActiveSheet.Paste
    
    For i = 2 To lastRowA
        lastRowH = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
        nextRowH = lastRowH + 1
        
        If i = 2 Then
            opening_price = ws.Cells(i, 3).Value
        End If
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ws.Select
            ws.Cells(i + 1, 1).Select
            Selection.Copy
            ws.Cells(nextRowH, "H").Select
            ActiveSheet.Paste
            closing_price = ws.Cells(i, 6).Value
            ws.Cells(nextRowH - 1, "I").Value = closing_price - opening_price
            ws.Cells(nextRowH - 1, "J").Value = (closing_price - opening_price) / opening_price
            opening_price = ws.Cells(i + 1, 3)
            If ws.Cells(nextRowH - 1, "I").Value > 0 Then
                ws.Cells(nextRowH - 1, "I").Interior.Color = vbGreen
            Else
               If ws.Cells(nextRowH - 1, "I").Value < 0 Then
                ws.Cells(nextRowH - 1, "I").Interior.Color = vbRed
                Else
                ws.Cells(nextRowH - 1, "I").Interior.Color = vbWhite
                End If
            End If
            
        End If

    

  Next i
    
  
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
        
        ws.Range("I2:I" & lastRowH).NumberFormat = "0.00"
        ws.Range("J2:J" & lastRowH).NumberFormat = "0.00%"
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).NumberFormat = "0.00%"



    
    Next ws

    MsgBox ("Complete")

End Sub

