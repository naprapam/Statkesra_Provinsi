Sub SisipkanTabelExcel_AppendAkhir()
    Dim xlApp As Object, xlWb As Object, ws As Object
    Dim dlgFile As FileDialog
    Dim lastRow As Long, lastCol As Long
    Dim rngExcel As Object
    Dim rngWord As Range
    Dim startTime As Double
    Dim gabung As String
    Dim i As Long
    Dim tbl As Table
    Dim tryCount As Integer
    
    ' === Pilih file Excel ===
    Set dlgFile = Application.FileDialog(msoFileDialogFilePicker)
    With dlgFile
        .Title = "Pilih file Excel Publikasi"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        If .Show <> -1 Then Exit Sub
    End With
    
    ' === Buka Excel ===
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlWb = xlApp.Workbooks.Open(dlgFile.SelectedItems(1))
    
    startTime = Timer
    Application.ScreenUpdating = False
    xlApp.ScreenUpdating = False
    
    ' === Loop semua sheet Excel ===
    For Each ws In xlWb.Sheets
        
        ' Skip sheet kosong
        If IsSheetEmpty(ws) Then GoTo NextSheet
        
        ' Tentukan lastRow & lastCol
        lastRow = ws.Cells(ws.Rows.Count, 1).End(-4161).Row   ' xlUp
        lastCol = ws.Cells(10, ws.Columns.Count).End(-4159).Column ' xlToLeft baris 10
        
        If lastRow < 1 Or LenB(Trim(CStr(ws.Cells(lastRow, 1).Value))) = 0 Then
            lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        End If
        If lastCol < 1 Then
            lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
        End If
        
        ' === Merge baris 1 dan 2 mulai dari kolom D sampai kolom terakhir ===
        If lastCol >= 4 Then
            ' Baris 1 langsung merge (tetap rata kiri, font asli)
            ws.Range(ws.Cells(1, 4), ws.Cells(1, lastCol)).Merge

	' Baris terakhir langsung merge (tetap rata kiri, font asli)
            ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, lastCol)).Merge
            
            ' Baris 2: gabungkan isi dulu ke D2
            gabung = ws.Cells(2, 4).Value
            For i = 5 To lastCol
                If Len(ws.Cells(2, i).Value) > 0 Then
                    gabung = gabung & " " & ws.Cells(2, i).Value
                End If
            Next i
            ws.Range(ws.Cells(2, 4), ws.Cells(2, lastCol)).Merge
            ws.Cells(2, 4).Value = gabung
            ' Alignment/font tidak diubah ? tetap ikut Excel asli
        End If
        
        ' Ambil range untuk copy
        Set rngExcel = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        
        ' === Paste di akhir dokumen ===
        rngExcel.Copy
        DoEvents
        
        Set rngWord = ActiveDocument.Content
        rngWord.Collapse Direction:=wdCollapseEnd
        
        ' Coba paste dengan retry
        tryCount = 0
        Do
            On Error Resume Next
            rngWord.PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False
            If Err.Number = 0 Then Exit Do
            Err.Clear
            DoEvents
            tryCount = tryCount + 1
        Loop While tryCount < 3
        On Error GoTo 0
        
        ' Jika tetap gagal, fallback paste biasa
        If tryCount >= 3 Then
            rngWord.Paste
        End If
        
        ' Format tabel terakhir
        If ActiveDocument.Tables.Count > 0 Then
            Set tbl = ActiveDocument.Tables(ActiveDocument.Tables.Count)
            tbl.AutoFitBehavior wdAutoFitContent
            tbl.PreferredWidthType = wdPreferredWidthPercent
            tbl.PreferredWidth = 100
            tbl.Rows.Alignment = wdAlignRowCenter ' center di halaman
	 tbl.AutoFitBehavior wdAutoFitWindow
        End If
        
        rngWord.InsertParagraphAfter
        
        ' === Page break antar sheet ===
        With ActiveDocument.Content
            .Collapse Direction:=wdCollapseEnd
            .InsertBreak Type:=wdPageBreak
        End With
        
NextSheet:
    Next ws
    
    ' === Tutup Excel ===
    xlWb.Close False
    xlApp.Quit
    
    Application.ScreenUpdating = True
    MsgBox "Selesai! Semua tabel ditempel di akhir dokumen (tiap sheet di halaman baru, tabel center, auto-fit halaman)."
End Sub

' Helper: deteksi sheet kosong
Private Function IsSheetEmpty(ws As Object) As Boolean
    Dim ur As Object
    Set ur = ws.UsedRange
    If ur Is Nothing Then
        IsSheetEmpty = True
    Else
        If ur.Cells.Count = 1 Then
            IsSheetEmpty = (Trim(CStr(ur.Value)) = "")
        Else
            IsSheetEmpty = False
        End If
    End If
End Function
