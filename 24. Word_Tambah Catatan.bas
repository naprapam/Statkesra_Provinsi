Sub IsiCatatanSuperscriptLurus_All_NoBorder_Strict()
    Dim rng As Range
    Dim t As Table
    Dim c As Cell
    Dim r As Row
    Dim rowIdx As Long
    
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "Kalimantan Selatan"
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            If rng.Information(wdWithInTable) Then
                rng.Select
                Selection.MoveDown Unit:=wdLine, Count:=1
                
                ' Format paragraf catatan
                With Selection.ParagraphFormat
                    .LeftIndent = 0
                    .FirstLineIndent = 0
                    .TabStops.ClearAll
                    .TabStops.Add Position:=CentimetersToPoints(3)
                End With
                
                ' --- Baris pertama catatan ---
                Selection.TypeText "Catatan/Note:" & vbTab
                Selection.TypeText "1"
                Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=True
                Selection.Font.Superscript = True
                Selection.Collapse wdCollapseEnd
                Selection.Font.Superscript = False
                Selection.TypeText " Jika RSE >25% tetapi ≤50%, estimasi harus digunakan dengan hati-hati/If RSE >25% but ≤50%, estimate should be used with caution."
                Selection.TypeParagraph
                
                ' --- Baris kedua catatan ---
                Selection.TypeText vbTab
                Selection.TypeText "2"
                Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=True
                Selection.Font.Superscript = True
                Selection.Collapse wdCollapseEnd
                Selection.Font.Superscript = False
                Selection.TypeText " Jika RSE >50%, estimasi dianggap tidak akurat/If RSE >50%, estimate considered unreliable."
                
                ' --- Hapus border bawah: target sel dan baris ---
                On Error Resume Next
                Set t = Selection.Tables(1)
                Set c = Selection.Cells(1)
                rowIdx = c.RowIndex
                Set r = t.Rows(rowIdx)
                
                ' 1) Hapus border bawah di sel catatan
                c.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                c.Borders(wdBorderBottom).LineWidth = wdLineWidth025pt  ' pastikan tidak tersisa
                
                ' 2) Hapus border bawah di baris catatan (kalau border datang dari baris)
                r.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                r.Borders(wdBorderBottom).LineWidth = wdLineWidth025pt
                
                ' 3) Hapus horizontal internal di baris atas/bawah yang kadang terlihat sebagai "garis bawah"
                If rowIdx > 1 Then t.Rows(rowIdx - 1).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                If rowIdx < t.Rows.Count Then t.Rows(rowIdx).Borders(wdBorderTop).LineStyle = wdLineStyleNone
                
                ' 4) Override style table supaya border tidak balik lagi
                t.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
                t.Rows(rowIdx).AllowBreakAcrossPages = False  ' tidak pengaruh ke border, tapi stabilkan baris
                On Error GoTo 0
            End If
            rng.Collapse wdCollapseEnd
        Loop
    End With
     
    MsgBox "Selesai: Catatan ditambahkan dan border bawah dihapus secara tegas.", vbInformation
End Sub
