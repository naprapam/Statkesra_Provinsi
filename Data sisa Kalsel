Sub Sisakan21BarisDenganKataKunci()
    Dim ws As Worksheet
    Dim rng As Range
    Dim barisAwal As Long, barisAkhir As Long
    
    For Each ws In ThisWorkbook.Sheets
        'Cari kata kunci di kolom A (bisa disesuaikan)
        Set rng = ws.Columns(1).Find(What:="R101 = Kalimantan Selatan", LookAt:=xlPart, LookIn:=xlValues)
        
        If Not rng Is Nothing Then
            barisAwal = rng.Row
            barisAkhir = barisAwal + 20   '21 baris total
            
            With ws
                Application.DisplayAlerts = False
                'Hapus semua baris di atas barisAwal
                If barisAwal > 1 Then
                    .Rows("1:" & barisAwal - 1).Delete
                End If
                'Karena baris sudah bergeser ke atas, barisAwal jadi 1
                barisAkhir = 21
                'Hapus semua baris setelah baris ke-21
                If barisAkhir < .Rows.Count Then
                    .Rows(barisAkhir + 1 & ":" & .Rows.Count).Delete
                End If
                Application.DisplayAlerts = True
            End With
        End If
    Next ws
End Sub
