Sub HapusKolomA_ClearBeforeMerge()
    Dim ws As Worksheet
    Dim mArea As Range, newRange As Range
    Dim val As Variant
    Dim valRow2 As Variant, valRow3 As Variant
    Dim startCol As Long, endCol As Long
    
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Worksheets
        ' 1. Simpan isi baris 2–3
        valRow2 = ws.Cells(2, 1).Value
        valRow3 = ws.Cells(3, 1).Value
        
        ' 2. Loop merge area hanya sekali (sel pertama)
        For Each mArea In ws.UsedRange
            If mArea.MergeCells Then
                If mArea.Address = mArea.MergeArea.Cells(1, 1).Address Then
                    ' Apakah merge area dimulai di kolom A?
                    If mArea.MergeArea.Column = 1 And mArea.Row <> 2 And mArea.Row <> 3 Then
                        val = mArea.Value
                        ' Hitung target: geser ke kanan, tapi lebar dikurangi 1
                        startCol = 2
                        endCol = mArea.MergeArea.Columns.Count
                        
                        mArea.UnMerge
                        
                        Set newRange = ws.Range(ws.Cells(mArea.Row, startCol), ws.Cells(mArea.Row, endCol))
                        If newRange.MergeCells Then newRange.UnMerge
                        
                        ' Kosongkan dulu target supaya tidak ada warning
                        newRange.ClearContents
                        
                        ' Merge ulang dan isi value
                        newRange.Merge
                        newRange.Value = val
                        newRange.HorizontalAlignment = xlCenter
                        newRange.VerticalAlignment = xlCenter
                        
                        ' Debug log
                        Debug.Print "Sheet: " & ws.Name & _
                                    " | Merge lama: " & mArea.MergeArea.Address & _
                                    " | Dipindah ke: " & newRange.Address & _
                                    " | Value: " & val
                    End If
                End If
            End If
        Next mArea
        
        ' 3. Hapus kolom A
        ws.Columns(1).Delete
        
        ' 4. Pulihkan baris 2–3 ke kolom A baru, rata kiri
        ws.Cells(2, 1).Value = valRow2
        ws.Cells(3, 1).Value = valRow3
        ws.Cells(2, 1).HorizontalAlignment = xlLeft
        ws.Cells(3, 1).HorizontalAlignment = xlLeft
    Next ws
    
    Application.ScreenUpdating = True
    MsgBox "Selesai! Kolom A dihapus, baris 2–3 tetap rata kiri, merge lain dipindah tanpa warning."
End Sub
