Sub GantiTeksPertamaKolomC_DenganNamaSheet()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
        
        ' loop dari atas ke bawah, cari teks pertama di kolom C
        For r = 1 To lastRow
            If Not IsEmpty(ws.Cells(r, "C").Value) Then
                If Not IsNumeric(ws.Cells(r, "C").Value) Then
                    ws.Cells(r, "C").Value = ws.Name
                    Exit For    ' berhenti setelah ketemu teks pertama
                End If
            End If
        Next r
    Next ws
End Sub
