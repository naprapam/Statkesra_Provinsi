Sub ValidasiPrediksiRSE_Final_Safe()
    Dim ws As Worksheet
    Dim rngLast As Range
    Dim lastRow As Long, r As Long
    Dim rngEst As Range, firstEst As Range
    Dim colEst As Long, colRSE As Long
    Dim rawPred As Variant, rawRSE As Variant
    Dim numPred As Double, numRSE As Double
    Dim okPred As Boolean, okRSE As Boolean, isDiv0 As Boolean
    
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
            
            For r = 2 To lastRow
                rawPred = ws.Cells(r, colEst).Value2
                rawRSE = ws.Cells(r, colRSE).Value2
                
                numPred = CleanPred(rawPred, okPred)
                numRSE = CleanRSE(rawRSE, isDiv0, okRSE)
                
                If isDiv0 Then
                    ws.Cells(r, colEst).Value = "-"
                ElseIf okPred And okRSE Then
                    If numRSE > 50 Then
                        ws.Cells(r, colEst).Value = "NA+="
                    ElseIf numRSE > 25 Then
                        ' Bulatkan 2 angka di belakang koma
                        ws.Cells(r, colEst).Value = Format(Round(numPred, 2), "0.00") & "=+"
                    Else
                        ws.Cells(r, colEst).Value = numPred
                    End If
                Else
                    ' Skip kalau salah satu bukan numeric dan bukan #DIV/0!
                End If
            Next r
            
            Set rngEst = ws.Cells.Find("Estimate", After:=rngEst, LookAt:=xlWhole, _
                                       MatchCase:=False, SearchDirection:=xlNext)
        Loop While Not rngEst Is Nothing And rngEst.Address <> firstEst.Address
        
NextSheet:
    Next ws
    
    MsgBox "Proses validasi selesai. Lihat Immediate Window (Ctrl+G).", vbInformation
End Sub
Private Function CleanRSE(ByVal v As Variant, ByRef isDiv0 As Boolean, ByRef isNumericOut As Boolean) As Double
    ' Normalize RSE: detect #DIV/0!, coerce numeric text/percent to Double
    isDiv0 = False
    isNumericOut = False
    CleanRSE = 0#
    
    If IsError(v) Then
        If v = CVErr(xlErrDiv0) Then
            isDiv0 = True
        End If
        Exit Function
    End If
    
    If IsEmpty(v) Then Exit Function
    
    Dim s As String
    s = CStr(v)
    s = Trim$(s)
    If s = "" Then Exit Function
    
    ' Strip percent sign if present
    If Right$(s, 1) = "%" Then s = Left$(s, Len(s) - 1)
    
    ' Replace locale separators if needed (comma decimal)
    ' Example: "12,5" ? "12.5"
    If InStr(s, ",") > 0 And InStr(s, ".") = 0 Then
        s = Replace(s, " ", "")
        s = Replace(s, ",", ".")
    End If
    
    ' Try converting
    On Error Resume Next
    Dim d As Double
    d = CDbl(s)
    If Err.Number = 0 Then
        isNumericOut = True
        CleanRSE = d
    End If
    On Error GoTo 0
End Function

Private Function CleanPred(ByVal v As Variant, ByRef isNumericOut As Boolean) As Double
    ' Normalize Prediksi: accept numeric or numeric text
    isNumericOut = False
    CleanPred = 0#
    
    If IsError(v) Or IsEmpty(v) Then Exit Function
    
    Dim s As String
    s = Trim$(CStr(v))
    If s = "" Then Exit Function
    
    On Error Resume Next
    Dim d As Double
    d = CDbl(s)
    If Err.Number = 0 Then
        isNumericOut = True
        CleanPred = d
    End If
    On Error GoTo 0
End Function

