Sub CopyAndModify_MultiSheets_Final()
    Dim wb As Workbook
    Dim ws As Worksheet, wsNew As Worksheet
    Dim lastCol As Long
    Dim colNum2 As Long, colNum3 As Long, colNum4 As Long, colNum5 As Long
    Dim sheetName As String, angkaRef As Variant
    Dim i As Long, totalSheets As Long
    
    Set wb = ThisWorkbook
    totalSheets = wb.Sheets.Count
    
    ' Loop mundur supaya sheet hasil copy tidak ikut diproses
    For i = totalSheets To 1 Step -1
        Set ws = wb.Sheets(i)
        sheetName = ws.Name
        
        ' Skip kalau sheet ini hasil copy
        If Right(sheetName, 2) = "_1" Or Right(sheetName, 2) = "_2" Then GoTo NextSheet
        
        ' Ambil angka referensi
        If IsError(ws.Cells(2, 3).Value) Then
            angkaRef = ""
        Else
            angkaRef = ws.Cells(2, 3).Value
        End If
        
        lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column
        colNum2 = FindColumnWithValue(ws.Rows(4), 2)
        colNum3 = FindColumnWithValue(ws.Rows(4), 3)
        colNum4 = FindColumnWithValue(ws.Rows(4), 4)
        colNum5 = FindColumnWithValue(ws.Rows(4), 5)
        
        Debug.Print "Proses sheet: " & sheetName
        Debug.Print "  colNum2=" & colNum2 & ", colNum3=" & colNum3 & ", colNum4=" & colNum4 & ", colNum5=" & colNum5 & ", lastCol=" & lastCol
        
        '=== Copy_1 ===
        If colNum2 > 0 And colNum3 > colNum2 Then
            ws.Copy After:=ws
            Set wsNew = ws.Next
            wsNew.Name = Left(sheetName & "_1", 31)
            Debug.Print "  -> Buat Copy_1: " & wsNew.Name
            
            If colNum2 >= 4 Then
                Debug.Print "     Hapus kolom D sampai " & colNum2
                wsNew.Range(wsNew.Cells(1, 4), wsNew.Cells(1, colNum2)).EntireColumn.Delete
            End If
            
            ' Tambahan: hapus semua kolom setelah angka 4
            Dim colNum4_copy As Long, lastCol_copy As Long
            lastCol_copy = wsNew.Cells(4, wsNew.Columns.Count).End(xlToLeft).Column
            colNum4_copy = FindColumnWithValue(wsNew.Rows(4), 4)
            If colNum4_copy > 0 And colNum4_copy < lastCol_copy Then
                Debug.Print "     Hapus semua kolom setelah angka 4 (col " & colNum4_copy & ")"
                wsNew.Range(wsNew.Cells(1, colNum4_copy + 1), wsNew.Cells(1, lastCol_copy)).EntireColumn.Delete
            End If
            
            Call ReplaceCaption(wsNew, angkaRef)
        End If
        
        '=== Copy_2 ===
        If colNum4 > 0 And colNum5 > colNum4 Then
            If SheetExists(sheetName & "_1") Then
                Set wsNew = wb.Sheets(sheetName & "_1")
                ws.Copy After:=wsNew
                Set wsNew = wsNew.Next
            Else
                ws.Copy After:=ws
                Set wsNew = ws.Next
            End If
            
            wsNew.Name = Left(sheetName & "_2", 31)
            Debug.Print "  -> Buat Copy_2: " & wsNew.Name
            
            If colNum4 >= 4 Then
                Debug.Print "     Hapus kolom D sampai " & colNum4
                wsNew.Range(wsNew.Cells(1, 4), wsNew.Cells(1, colNum4)).EntireColumn.Delete
            End If
            
            Call ReplaceCaption(wsNew, angkaRef)
        End If
        
        '=== Potong sheet asal terakhir ===
        If colNum2 > 0 And colNum3 > colNum2 Then
            Debug.Print "  -> Potong sheet asal setelah kolom " & colNum2
            ws.Range(ws.Cells(1, colNum2 + 1), ws.Cells(1, lastCol)).EntireColumn.Delete
        End If
        
NextSheet:
    Next i
    
    MsgBox "Proses multi-sheet selesai! Lihat Immediate Window (Ctrl+G) untuk log."
End Sub

Function FindColumnWithValue(rng As Range, val As Variant) As Long
    Dim c As Range
    For Each c In rng.Cells
        If Not IsError(c.Value) Then
            If IsNumeric(c.Value) Then
                If CLng(c.Value) = val Then
                    FindColumnWithValue = c.Column
                    Exit Function
                End If
            End If
        End If
    Next c
    FindColumnWithValue = 0
End Function

Function SheetExists(shtName As String) As Boolean
    Dim sht As Worksheet
    On Error Resume Next
    Set sht = ThisWorkbook.Sheets(shtName)
    SheetExists = Not sht Is Nothing
    On Error GoTo 0
End Function

Sub ReplaceCaption(ws As Worksheet, angkaRef As Variant)
    Dim rng As Range, c As Range
    Dim newText As String, posSlash As Long, posNum As Long
    Dim lastCol As Long, angkaText As String
    
    lastCol = ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column
    angkaText = CStr(angkaRef)
    
    Set rng = ws.UsedRange
    For Each c In rng
        If InStr(1, c.Value, "Tabel", vbTextCompare) > 0 Then
            newText = "Lanjutan Tabel/Continued Table " & angkaText
            c.Value = newText
            
            posSlash = InStr(newText, "/")
            posNum = InStrRev(newText, " ") + 1
            
            c.Font.Bold = False
            c.Font.Italic = False
            
            c.Characters(1, posSlash - 1).Font.Bold = True
            c.Characters(posSlash, Len("Continued Table") + 1).Font.Italic = True
            c.Characters(posNum, Len(angkaText)).Font.Bold = True
            
            
            '=== Merge caption sampai kolom terakhir, rata kiri ===
            ws.Range(ws.Cells(c.Row, 1), ws.Cells(c.Row, lastCol)).Merge
            ws.Cells(c.Row, 1).HorizontalAlignment = xlLeft
            ws.Cells(c.Row, 1).VerticalAlignment = xlCenter
        End If
    Next c
End Sub
