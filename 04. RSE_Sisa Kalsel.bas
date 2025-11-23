Sub RapikanSemuaSheet_Kalsel()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim kalSelRow As Long, startKeep As Long
    Dim keepRows As Collection
    
    Application.ScreenUpdating = False
    
    ' Loop semua sheet di workbook aktif
    For Each ws In ThisWorkbook.Sheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Cari baris terakhir yang mengandung "Kalimantan Selatan"
        kalSelRow = 0
        For r = lastRow To 1 Step -1
            If InStr(1, ws.Cells(r, "A").Value, "Kalimantan Selatan", vbTextCompare) > 0 Then
                kalSelRow = r
                Exit For
            End If
        Next r
        
        If kalSelRow > 0 Then
            ' Simpan baris yang dipertahankan
            Set keepRows = New Collection
            ' Baris 1–10
            For r = 1 To 10
                keepRows.Add r
            Next r
            
            ' 13 baris sebelum KalSel
            startKeep = kalSelRow - 13
            If startKeep < 11 Then startKeep = 11
            For r = startKeep To kalSelRow - 1
                keepRows.Add r
            Next r
            
            ' Baris KalSel
            keepRows.Add kalSelRow
            
            ' Hapus baris lain dari bawah ke atas
            For r = lastRow To 1 Step -1
                If Not IsInCollection(keepRows, r) Then
                    ws.Rows(r).Delete
                End If
            Next r
        End If
    Next ws
    
    Application.ScreenUpdating = True
    MsgBox "Selesai: semua sheet sudah dirapikan (baris 1–10, 13 baris sebelum KalSel, dan baris KalSel dipertahankan)."
End Sub

' Fungsi bantu untuk cek apakah nilai ada di Collection
Private Function IsInCollection(col As Collection, val As Long) As Boolean
    Dim itm As Variant
    For Each itm In col
        If itm = val Then
            IsInCollection = True
            Exit Function
        End If
    Next itm
End Function
