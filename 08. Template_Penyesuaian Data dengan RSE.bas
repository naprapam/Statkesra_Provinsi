Sub TandaiNA_RSE_LogEstimasi_WithRSE()
    Dim wbPub As Workbook, wbRef As Workbook
    Dim wsPub As Worksheet, wsRef As Worksheet, wsLog As Worksheet
    Dim FileRef As String
    Dim lastRowPub As Long, lastColPub As Long
    Dim r As Long, c As Long
    Dim namaSheetRef As String
    Dim rngCari As Range, cellRef As Range
    Dim valRSE As Variant, valPub As Variant, valRef As Variant
    Dim namaDaerahPub As String, namaDaerahRef As String
    Dim logRow As Long, foundMatch As Boolean
    
    Set wbPub = ThisWorkbook
    
    ' Buat / reset sheet log estimasi
    On Error Resume Next
    Set wsLog = wbPub.Sheets("LogEstimasi")
    If wsLog Is Nothing Then
        Set wsLog = wbPub.Sheets.Add
        wsLog.Name = "LogEstimasi"
    End If
    wsLog.Cells.Clear
    wsLog.Range("A1:F1").Value = Array("Sheet", "Row", "Col", "PubValue", "RSE", "Keterangan")
    logRow = 2
    On Error GoTo 0
    
    ' Pilih file referensi
    FileRef = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*")
    If FileRef = "False" Then Exit Sub
    Set wbRef = Workbooks.Open(FileRef, ReadOnly:=False)
    
    Application.ScreenUpdating = False
    
    ' Loop semua sheet publikasi
    For Each wsPub In wbPub.Sheets
        If wsPub.Name <> "LogEstimasi" Then
            namaSheetRef = "9." & wsPub.Name
            On Error Resume Next
            Set wsRef = wbRef.Sheets(namaSheetRef)
            On Error GoTo 0
            
            If Not wsRef Is Nothing Then
                lastRowPub = wsPub.Cells(wsPub.Rows.Count, "C").End(xlUp).Row
                lastColPub = wsPub.Cells(6, wsPub.Columns.Count).End(xlToLeft).Column
                
                ' Loop baris publikasi
                For r = 1 To lastRowPub
                    namaDaerahPub = LCase(Trim(wsPub.Cells(r, "C").Value))
                    If Len(namaDaerahPub) > 0 Then
                        ' Cari nama daerah di referensi (kolom B)
                        For Each cellRef In wsRef.Range("B1:B" & wsRef.Cells(wsRef.Rows.Count, "B").End(xlUp).Row)
                            namaDaerahRef = LCase(Trim(Mid(cellRef.Value, 7)))
                            
                            ' Normalisasi khusus
                            If InStr(namaDaerahRef, "banjar baru") > 0 Then
                                namaDaerahRef = Replace(namaDaerahRef, "banjar baru", "banjarbaru")
                            End If
                            If InStr(namaDaerahRef, "kota baru") > 0 Then
                                namaDaerahRef = Replace(namaDaerahRef, "kota baru", "kotabaru")
                            End If
                            
                            If namaDaerahRef = namaDaerahPub Then
                                ' Loop semua kolom angka di publikasi
                                For c = 4 To lastColPub
                                    valPub = wsPub.Cells(r, c).Value
                                    If IsNumeric(valPub) And valPub <> "" Then
                                        valPub = Round(CDbl(valPub), 2)
                                        foundMatch = False
                                        
                                        ' Cari nilai estimasi yang sama di baris referensi
                                        For Each rngCari In wsRef.Rows(cellRef.Row).Cells
                                            If IsNumeric(rngCari.Value) Then
                                                valRef = Round(CDbl(rngCari.Value), 2)
                                                If valRef = valPub Then
                                                    foundMatch = True
                                                    
                                                    ' Ambil RSE dari kolom +2
                                                    valRSE = wsRef.Cells(cellRef.Row, rngCari.Column + 2).Value
                                                    
                                                    ' Jika referensi = 0.00 → ganti "–"
                                                    If valRef = 0 Then
                                                        wsPub.Cells(r, c).Value = "–"
                                                        wsLog.Cells(logRow, 1).Value = wsPub.Name
                                                        wsLog.Cells(logRow, 2).Value = r
                                                        wsLog.Cells(logRow, 3).Value = c
                                                        wsLog.Cells(logRow, 4).Value = valPub
                                                        wsLog.Cells(logRow, 5).Value = valRSE
                                                        wsLog.Cells(logRow, 6).Value = "Referensi 0.00 → diganti –"
                                                        logRow = logRow + 1
                                                    
                                                    ElseIf IsNumeric(valRSE) Then
                                                        If CDbl(valRSE) > 50 Then
                                                            wsPub.Cells(r, c).Value = "NA"
                                                            wsLog.Cells(logRow, 1).Value = wsPub.Name
                                                            wsLog.Cells(logRow, 2).Value = r
                                                            wsLog.Cells(logRow, 3).Value = c
                                                            wsLog.Cells(logRow, 4).Value = valPub
                                                            wsLog.Cells(logRow, 5).Value = valRSE
                                                            wsLog.Cells(logRow, 6).Value = "RSE > 50 → diganti NA"
                                                            logRow = logRow + 1
                                                        End If
                                                    End If
                                                    Exit For
                                                End If
                                            End If
                                        Next rngCari
                                        
                                        ' Kalau tidak ketemu sama sekali → log
                                        If Not foundMatch Then
                                            wsLog.Cells(logRow, 1).Value = wsPub.Name
                                            wsLog.Cells(logRow, 2).Value = r
                                            wsLog.Cells(logRow, 3).Value = c
                                            wsLog.Cells(logRow, 4).Value = valPub
                                            wsLog.Cells(logRow, 5).Value = ""   ' RSE kosong
                                            wsLog.Cells(logRow, 6).Value = "Estimasi tidak ketemu di referensi"
                                            logRow = logRow + 1
                                        End If
                                    End If
                                Next c
                                Exit For ' keluar setelah ketemu baris cocok
                            End If
                        Next cellRef
                    End If
                Next r
            End If
            Set wsRef = Nothing
        End If
    Next wsPub
    
    Application.ScreenUpdating = True
    wbRef.Close SaveChanges:=True
    
    MsgBox "Selesai: cek selesai. Lihat sheet 'LogEstimasi' untuk daftar NA, – , dan estimasi yang tidak ketemu (dengan RSE)."
End Sub
