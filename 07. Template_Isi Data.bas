Sub CopyDataSumberKeTemplate_KodeWilayah()
    Dim wbSumber As Workbook, wbTemplate As Workbook
    Dim wsTemplate As Worksheet, wsSumber As Worksheet
    Dim fd As FileDialog, FilePath As String
    Dim nomorTabel As String, namaSheetSumber As String
    Dim lastRowSumber As Long, lastColSumber As Long
    Dim lastRowTemplate As Long, r As Long
    Dim startRow As Long
    Dim dataArr As Variant, i As Long, j As Long
    
    ' Pilih file sumber
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Pilih File Sumber Data (sheet bernama tabel 1, tabel 2, dst)"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        If .Show <> -1 Then Exit Sub
        FilePath = .SelectedItems(1)
    End With
    
    ' Buka file sumber
    Set wbSumber = Workbooks.Open(FilePath, ReadOnly:=True)
    Set wbTemplate = ThisWorkbook
    
    Application.ScreenUpdating = False
    
    ' Loop semua sheet di template
    For Each wsTemplate In wbTemplate.Sheets
        lastRowTemplate = wsTemplate.Cells(wsTemplate.Rows.Count, "C").End(xlUp).Row
        
        ' Loop semua baris di kolom C (nomor tabel)
        For r = 1 To lastRowTemplate
            If IsNumeric(wsTemplate.Cells(r, "C").Value) Then
                nomorTabel = Trim(CStr(wsTemplate.Cells(r, "C").Value))
                namaSheetSumber = "tabel " & nomorTabel
                
                ' Cek apakah sheet sumber ada
                On Error Resume Next
                Set wsSumber = wbSumber.Sheets(namaSheetSumber)
                On Error GoTo 0
                
                If Not wsSumber Is Nothing Then
                    ' Tentukan baris awal data berdasarkan A6/A7
                    If LCase(Trim(wsSumber.Cells(6, "A").Value)) Like "6301*" Then
                        startRow = 6
                    ElseIf LCase(Trim(wsSumber.Cells(7, "A").Value)) Like "6301*" Then
                        startRow = 7
                    Else
                        ' fallback: cari baris pertama berisi kode wilayah
                        startRow = wsSumber.Columns("A").Find(What:="6301*", LookAt:=xlPart, LookIn:=xlValues).Row
                    End If
                    
                    ' Cari ukuran tabel di sumber
                    lastRowSumber = wsSumber.Cells(wsSumber.Rows.Count, "B").End(xlUp).Row
                    lastColSumber = wsSumber.Cells(startRow, wsSumber.Columns.Count).End(xlToLeft).Column
                    
                    ' Ambil data ke array (mulai kolom B)
                    dataArr = wsSumber.Range(wsSumber.Cells(startRow, 2), _
                                             wsSumber.Cells(lastRowSumber, lastColSumber)).Value
                    
                    ' Bulatkan ke 2 angka desimal
                    For i = 1 To UBound(dataArr, 1)
                        For j = 1 To UBound(dataArr, 2)
                            If IsNumeric(dataArr(i, j)) Then
                                dataArr(i, j) = Round(CDbl(dataArr(i, j)), 2)
                            End If
                        Next j
                    Next i
                    
                   ' Cari baris "Tanah Laut" di template
                        Set rngPaste = wsTemplate.Columns("C").Find(What:="Tanah Laut", LookAt:=xlPart, LookIn:=xlValues)
                        If Not rngPaste Is Nothing Then
                            pasteRow = rngPaste.Row
                            wsTemplate.Cells(pasteRow, "E").Resize(UBound(dataArr, 1), UBound(dataArr, 2)).Value = dataArr
                    End If
                End If
                Set wsSumber = Nothing
            End If
        Next r
    Next wsTemplate
    
    Application.ScreenUpdating = True
    wbSumber.Close SaveChanges:=False
    
    MsgBox "Selesai: data dari sheet sumber sudah ditempel ke kolom E mulai baris 7 dengan pembulatan 2 desimal."
End Sub
