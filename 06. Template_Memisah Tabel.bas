Sub PisahkanTabel_MultiSheet_FormatLengkap()
    Dim wsSumber As Worksheet, wsBaru As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, startRow As Long, endRow As Long
    Dim namaSheet As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Loop semua sheet di workbook
    For Each wsSumber In ThisWorkbook.Sheets
        lastRow = wsSumber.Cells(wsSumber.Rows.Count, "A").End(xlUp).Row
        startRow = 0
        ' Tentukan batas kolom dari baris 6
        lastCol = wsSumber.Cells(6, wsSumber.Columns.Count).End(xlToLeft).Column
        
        For r = 1 To lastRow
            ' Marker: baris dengan kolom A berisi "Tabel"
            If LCase(Trim(wsSumber.Cells(r, "A").Value)) Like "tabel*" Then
                ' Jika sebelumnya ada tabel, selesaikan dulu
                If startRow > 0 Then
                    endRow = r - 1
                    
                    ' Nama sheet dari kolom C baris marker
                    namaSheet = Trim(wsSumber.Cells(startRow, "C").Value)
                    
                    ' Hapus sheet lama jika sudah ada
                    On Error Resume Next
                    ThisWorkbook.Sheets(namaSheet).Delete
                    On Error GoTo 0
                    
                    ' Buat sheet baru
                    Set wsBaru = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                    wsBaru.Name = namaSheet
                    
                    ' Copy blok tabel dengan format, lebar kolom, dan tinggi baris
                    wsSumber.Range(wsSumber.Cells(startRow, 1), wsSumber.Cells(endRow, lastCol)).Copy
                    With wsBaru.Range("A1")
                        .PasteSpecial Paste:=xlPasteAll
                        .PasteSpecial Paste:=xlPasteColumnWidths
                    End With
                    ' Samakan tinggi baris
                    wsBaru.Rows.RowHeight = wsSumber.Rows.RowHeight
                End If
                
                ' Set awal tabel baru
                startRow = r
            End If
        Next r
        
        ' Tangani tabel terakhir di sheet
        If startRow > 0 And startRow < lastRow Then
            endRow = lastRow
            namaSheet = Trim(wsSumber.Cells(startRow, "C").Value)
            
            On Error Resume Next
            ThisWorkbook.Sheets(namaSheet).Delete
            On Error GoTo 0
            
            Set wsBaru = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsBaru.Name = namaSheet
            
            wsSumber.Range(wsSumber.Cells(startRow, 1), wsSumber.Cells(endRow, lastCol)).Copy
            With wsBaru.Range("A1")
                .PasteSpecial Paste:=xlPasteAll
                .PasteSpecial Paste:=xlPasteColumnWidths
            End With
            wsBaru.Rows.RowHeight = wsSumber.Rows.RowHeight
        End If
    Next wsSumber
    
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Selesai: semua tabel dari semua sheet sudah dipisahkan ke sheet baru dengan format, lebar kolom, dan tinggi baris dipertahankan."
End Sub
