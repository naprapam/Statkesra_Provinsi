Sub FormatRegencyMunicipality_TopCenter()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        For r = 1 To lastRow
            For c = 1 To lastCol
                ' cek persis sama dengan "Regency/Municipality"
                If StrComp(Trim(ws.Cells(r, c).Value), "Regency/Municipality", vbTextCompare) = 0 Then
                    With ws.Cells(r, c)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlTop
		  .Font.Bold = False
		  .Font.Italic = True
                    End With
                End If
            Next c
        Next r
    Next ws
End Sub
