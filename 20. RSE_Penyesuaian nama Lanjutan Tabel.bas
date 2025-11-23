Sub FormatLanjutanContinuedTable_WithNumber()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim val As String
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        For r = 1 To lastRow
            For c = 1 To lastCol
                val = ws.Cells(r, c).Value
                
                ' kalau ada kata "Lanjutan Tabel"
                If InStr(1, val, "Lanjutan Tabel", vbTextCompare) > 0 Then
                    With ws.Cells(r, c)
                        .HorizontalAlignment = xlLeft
                        .VerticalAlignment = xlCenter
                        .Font.Bold = True
                    End With
                End If
                
                ' kalau ada kata "Continued Table"
                If InStr(1, val, "Continued Table", vbTextCompare) > 0 Then
                    With ws.Cells(r, c)
                        .HorizontalAlignment = xlLeft
                        .VerticalAlignment = xlCenter
                        .Font.Italic = True
                        .Font.Bold = False
                    End With
                End If
            Next c
        Next r
    Next ws
End Sub
