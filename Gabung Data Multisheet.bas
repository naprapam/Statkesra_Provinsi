Sub GabungkanFileExcel_MultiSheet()
    Dim wbTujuan As Workbook
    Dim wbSumber As Workbook
    Dim ws As Worksheet
    Dim FileDialog As FileDialog
    Dim FilePath As Variant
    Dim NamaFile As String
    Dim NamaSheetBaru As String
    
    Set wbTujuan = ThisWorkbook
    
    'Pilih beberapa file
    Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)
    With FileDialog
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        
        If .Show = -1 Then
            For Each FilePath In .SelectedItems
                'Ambil nama file tanpa path & ekstensi
                NamaFile = Mid(FilePath, InStrRev(FilePath, "\") + 1)
                NamaFile = Left(NamaFile, InStrRev(NamaFile, ".") - 1)
                
                Set wbSumber = Workbooks.Open(FilePath)
                
                'Loop semua sheet di file sumber
                For Each ws In wbSumber.Sheets
                    ws.Copy After:=wbTujuan.Sheets(wbTujuan.Sheets.Count)
                    
                    'Buat nama sheet unik: NamaFile_SheetName
                    NamaSheetBaru = NamaFile & "_" & ws.Name
                    
                    'Jika nama terlalu panjang atau sudah ada, tambahkan suffix
                    On Error Resume Next
                    wbTujuan.Sheets(wbTujuan.Sheets.Count).Name = Left(NamaSheetBaru, 31)
                    If Err.Number <> 0 Then
                        Err.Clear
                        wbTujuan.Sheets(wbTujuan.Sheets.Count).Name = Left(NamaFile, 25) & "_" & Format(Now, "hhmmss")
                    End If
                    On Error GoTo 0
                Next ws
                
                wbSumber.Close SaveChanges:=False
            Next FilePath
        End If
    End With
    
    MsgBox "Selesai: semua sheet dari file terpilih sudah digabung ke workbook ini."
End Sub
