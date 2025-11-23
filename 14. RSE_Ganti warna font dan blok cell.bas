Sub FormatKabupatenKota_Exact_AllSheets()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        For r = 1 To lastRow
            For c = 1 To lastCol
                ' cek persis sama dengan "Kabupaten/Kota" atau "Regency/Municipality"
                If StrComp(Trim(ws.Cells(r, c).Value), "Kabupaten/Kota", vbTextCompare) = 0 _
                   Or StrComp(Trim(ws.Cells(r, c).Value), "Regency/Municipality", vbTextCompare) = 0 Then
                   
                    With ws.Cells(r, c)
                        .Font.Color = RGB(255, 255, 255)          ' putih
                        .Interior.Color = RGB(55, 104, 145)       ' #376891
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlBottom
                        .Font.Bold = True
                    End With
                End If
            Next c
        Next r
    Next ws
End Sub
