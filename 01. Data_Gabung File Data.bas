Sub GabungkanFileExcel()
    Dim wbTujuan As Workbook
    Dim wbSumber As Workbook
    Dim ws As Worksheet
    Dim FileDialog As FileDialog
    Dim FilePath As Variant
    Dim NamaFile As String
    
    Set wbTujuan = ThisWorkbook
    
    'Pilih beberapa file
    Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)
    With FileDialog
        .AllowMultiSelect = True
        .Filters.Add "Excel Files", "*.xls; *.xlsx"
        If .Show = -1 Then
            For Each FilePath In .SelectedItems
                'Ambil nama file tanpa path & ekstensi
                NamaFile = Mid(FilePath, InStrRev(FilePath, "\") + 1)
                NamaFile = Left(NamaFile, InStrRev(NamaFile, ".") - 1)
                
                Set wbSumber = Workbooks.Open(FilePath)
                
                'Ambil hanya sheet pertama (atau bisa di-loop semua sheet)
                wbSumber.Sheets(1).Copy After:=wbTujuan.Sheets(wbTujuan.Sheets.Count)
                
                'Rename sheet hasil copy sesuai nama file
                wbTujuan.Sheets(wbTujuan.Sheets.Count).Name = NamaFile
                
                wbSumber.Close SaveChanges:=False
            Next FilePath
        End If
    End With
End Sub
