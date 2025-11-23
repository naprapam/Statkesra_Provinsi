Sub ValidasiPrediksiRSE_Final_Safe()
    Dim ws As Worksheet
    Dim rngLast As Range
    Dim lastRow As Long, r As Long
    Dim rngEst As Range, firstEst As Range
    Dim colEst As Long, colRSE As Long
    Dim colSE As Long, colBB As Long, colBA As Long
    Dim rawPred As Variant, rawRSE As Variant
    Dim numPred As Double, numRSE As Double
    Dim okPred As Boolean, okRSE As Boolean, isDiv0 As Boolean
    Dim predText As String
    
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print "=== Sheet: " & ws.Name & " ==="
        
        Set rngLast = ws.Columns(1).Find("Kalimantan Selatan", LookAt:=xlPart, MatchCase:=False)
        If rngLast Is Nothing Then GoTo NextSheet
        lastRow = rngLast.Row
        
        Set rngEst = ws.Cells.Find("Estimate", LookAt:=xlWhole, MatchCase:=False)
        If rngEst Is Nothing Then GoTo NextSheet
        Set firstEst = rngEst
        
        Do
            colEst = rngEst.Column
            colRSE = colEst + 2
            colSE = colEst + 1
            colBB = colEst + 3
            colBA = colEst + 4
            
            For r = 2 To lastRow
                rawPred = ws.Cells(r, colEst).Value2
                rawRSE = ws.Cells(r, colRSE).Value2
                
                numPred = CleanPred(rawPred, okPred)
                numRSE = CleanRSE(rawRSE, isDiv0, okRSE)
                
                ' --- Jika RSE error (#DIV/0!) ? semua jadi "-" ---
                If isDiv0 Then
                    ws.Cells(r, colEst).Value = "-"
                    ws.Cells(r, colRSE).Value = "-"
                    ws.Cells(r, colSE).Value = "-"
                    ws.Cells(r, colBB).Value = "-"
                    ws.Cells(r, colBA).Value = "-"
                
                Else
                    ' --- Bulatkan & format RSE (jika valid) ---
                    If okRSE Then
                        ws.Cells(r, colRSE).Value = Round(numRSE, 2)
                        ws.Cells(r, colRSE).NumberFormat = "0.00"
                    End If
                    
                    ' --- Aturan dasar Prediksi ---
                    If okPred And okRSE Then
                        If numRSE > 50 Then
                            ws.Cells(r, colEst).Value = "NA+="
                        ElseIf numRSE > 25 Then
                            ws.Cells(r, colEst).Value = Format(Round(numPred, 2), "0.00") & "=+"
                        Else
                            ws.Cells(r, colEst).Value = Format(Round(numPred, 2), "0.00")
                        End If
                    End If
                    
                    ' --- Aturan tambahan Prediksi ? RSE/SE/BB/BA ---
                    predText = CStr(ws.Cells(r, colEst).Value)
                    
                    If predText = "NA+=" Then
                        ' RSE diberi tanda =+, SE/BB/BA jadi "-"
                        ws.Cells(r, colRSE).Value = CStr(ws.Cells(r, colRSE).Value) & "+="
                        ws.Cells(r, colSE).Value = "-"
                        ws.Cells(r, colBB).Value = "-"
                        ws.Cells(r, colBA).Value = "-"
                    ElseIf InStr(predText, "=+") > 0 Then
                        ' RSE diberi tanda =+, SE/BB/BA tetap
                        ws.Cells(r, colRSE).Value = CStr(ws.Cells(r, colRSE).Value) & "=+"
                    End If
                End If
            Next r
            
            Set rngEst = ws.Cells.Find("Estimate", After:=rngEst, LookAt:=xlWhole, _
                                       MatchCase:=False, SearchDirection:=xlNext)
        Loop While Not rngEst Is Nothing And rngEst.Address <> firstEst.Address
        
NextSheet:
    Next ws
    
    MsgBox "Proses validasi selesai. Lihat Immediate Window (Ctrl+G).", vbInformation
End Sub

' === Helper Functions ===
Private Function CleanRSE(ByVal v As Variant, ByRef isDiv0 As Boolean, ByRef isNumericOut As Boolean) As Double
    isDiv0 = False
    isNumericOut = False
    CleanRSE = 0#
    
    If IsError(v) Then
        If v = CVErr(xlErrDiv0) Then isDiv0 = True
        Exit Function
    End If
    If IsEmpty(v) Then Exit Function
    
    Dim s As String, d As Double
    s = Trim$(CStr(v))
    If s = "" Then Exit Function
    
    If Right$(s, 1) = "%" Then s = Left$(s, Len(s) - 1)
    If InStr(s, ",") > 0 And InStr(s, ".") = 0 Then
        s = Replace(s, " ", "")
        s = Replace(s, ",", ".")
    End If
    
    On Error Resume Next
    d = CDbl(s)
    If Err.Number = 0 Then
        isNumericOut = True
        CleanRSE = d
    End If
    On Error GoTo 0
End Function

Private Function CleanPred(ByVal v As Variant, ByRef isNumericOut As Boolean) As Double
    isNumericOut = False
    CleanPred = 0#
    
    If IsError(v) Or IsEmpty(v) Then Exit Function
    
    Dim s As String, d As Double
    s = Trim$(CStr(v))
    If s = "" Then Exit Function
    
    On Error Resume Next
    d = CDbl(s)
    If Err.Number = 0 Then
        isNumericOut = True
        CleanPred = d
    End If
    On Error GoTo 0
End Function
