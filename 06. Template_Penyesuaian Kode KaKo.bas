Option Explicit

Private Function LookupNama(wsMaster As Worksheet, ByVal kode As Variant) As String
    Dim rng As Range, kStr As String
    kStr = Trim(CStr(kode))
    If IsNumeric(kStr) Then kStr = Format(Val(kStr), "00")
    Set rng = wsMaster.Columns(1).Find(What:=kStr, LookAt:=xlWhole, LookIn:=xlValues)
    If Not rng Is Nothing Then LookupNama = Trim(CStr(rng.Offset(0, 1).Value)): Exit Function
    Set rng = wsMaster.Columns(1).Find(What:=Trim(CStr(kode)), LookAt:=xlWhole, LookIn:=xlValues)
    If Not rng Is Nothing Then LookupNama = Trim(CStr(rng.Offset(0, 1).Value)): Exit Function
    If IsNumeric(kode) Then
        Set rng = wsMaster.Columns(1).Find(What:=Val(kode), LookAt:=xlWhole, LookIn:=xlValues)
        If Not rng Is Nothing Then LookupNama = Trim(CStr(rng.Offset(0, 1).Value)): Exit Function
    End If
    LookupNama = ""
End Function

Public Sub UpdateSemuaSheetDenganMaster_Adaptif()
    Dim wbMaster As Workbook, wsMaster As Worksheet
    Dim fd As FileDialog, FilePath As String
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, startBlk As Long, endBlk As Long
    Dim k As Variant, nama As String, i As Long, relRow As Long
    
    ' Pilih file master
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Pilih File Master Kabupaten/Kota"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        If .Show <> -1 Then Exit Sub
        FilePath = .SelectedItems(1)
    End With
    
    Set wbMaster = Workbooks.Open(FilePath, ReadOnly:=True)
    Set wsMaster = wbMaster.Sheets(1)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Visible = xlSheetVisible Then
            ws.Cells.Replace What:="[Nama Provinsi]", Replacement:="Kalimantan Selatan", LookAt:=xlPart
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            r = 1
            Do While r <= lastRow
                Do While r <= lastRow And Not IsNumeric(ws.Cells(r, "A").Value)
                    r = r + 1
                Loop
                If r > lastRow Then Exit Do
                startBlk = r
                endBlk = startBlk
                Do While endBlk <= lastRow And IsNumeric(ws.Cells(endBlk, "A").Value)
                    endBlk = endBlk + 1
                Loop
                endBlk = endBlk - 1
                If endBlk < startBlk Then Exit Do
                
                ' === 01–11 ===
                For i = startBlk + 1 To Application.WorksheetFunction.Min(endBlk, startBlk + 12)
                    k = ws.Cells(i, "A").Value
                    ws.Cells(i, "B").Value = "Kab."
                    nama = LookupNama(wsMaster, k)
                    If Len(nama) > 0 Then ws.Cells(i, "C").Value = nama
                Next i
                
                ' === Baris ke-12 relatif blok ? 71 Kota. ===
                relRow = startBlk + 12
                If relRow <= endBlk Then
                    ws.Cells(relRow, "A").Value = "71."
                    ws.Cells(relRow, "B").Value = "Kota."
                    nama = LookupNama(wsMaster, "71.")
                    If Len(nama) > 0 Then ws.Cells(relRow, "C").Value = nama
                End If
                
                ' === Baris ke-13 relatif blok ? 72 Kota. ===
                relRow = startBlk + 13
                If relRow <= endBlk Then
                    ws.Cells(relRow, "A").Value = "72."
                    ws.Cells(relRow, "B").Value = "Kota."
                    nama = LookupNama(wsMaster, "72.")
                    If Len(nama) > 0 Then ws.Cells(relRow, "C").Value = nama
                End If
                
                ' Hapus baris setelah ke-13 relatif blok
                Dim hapusMulai As Long
                hapusMulai = startBlk + 14
                If endBlk >= hapusMulai Then
                    ws.Rows(hapusMulai & ":" & endBlk).Delete
                    endBlk = startBlk + 13
                    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                End If

	     ' === Tambahan: Kolom C rata kiri untuk blok ini ===
	     ws.Range(ws.Cells(startBlk, "C"), ws.Cells(endBlk, "C")).HorizontalAlignment = xlLeft
	                
                r = endBlk + 1
            Loop
        End If
    Next ws
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    wbMaster.Close SaveChanges:=False
    MsgBox "Selesai: semua sheet diproses, kolom C terisi, Kab. mulai 1 baris di bawah header, 72 = Kota.", vbInformation
End Sub
