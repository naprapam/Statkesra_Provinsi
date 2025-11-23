Sub MergeKabupatenKota_AllSheets()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim colKabKota As Long
    Dim rngMerge As Range
    
    ' loop semua sheet
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        For i = 1 To lastRow - 1
            ' deteksi cell berisi "Kabupaten/Kota"
            If InStr(1, ws.Cells(i, 1).Value, "Kabupaten/Kota", vbTextCompare) > 0 Then
                colKabKota = ws.Cells(i, 1).Column
                
                ' merge cell di kolom Kabupaten/Kota dengan baris bawahnya
                Set rngMerge = ws.Range(ws.Cells(i, colKabKota), ws.Cells(i + 1, colKabKota))
                rngMerge.Merge
                
                ' format hasil merge: center + bottom align
                With rngMerge
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                End With
            End If
        Next i
    Next ws
End Sub
