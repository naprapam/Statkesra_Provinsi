Sub NormalisasiAngkaKisaran100()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim rawText As String
    Dim numPart As String
    Dim suffix As String
    Dim numVal As Double
    Dim i As Long
    
    For Each ws In ThisWorkbook.Worksheets
        Set rng = ws.UsedRange
        For Each cell In rng
            If Not IsEmpty(cell.Value) Then
                rawText = CStr(cell.Value)
                
                ' Cari bagian angka di depan
                i = 1
                Do While i <= Len(rawText) And (Mid(rawText, i, 1) Like "[0-9.,E+-]")
                    i = i + 1
                Loop
                
                numPart = Left(rawText, i - 1)   ' angka di depan
                suffix = Mid(rawText, i)        ' huruf/simbol di belakang
                
                If numPart <> "" And IsNumeric(numPart) Then
                    numVal = Val(numPart)
                    
                    ' Aturan pembulatan
                    If numVal >= 100 And numVal < 200 Then
                        cell.Value = Format(100, "0.00") & suffix
                    ElseIf numVal >= 0 And numVal < 0.01 Then
                        cell.Value = "~0" & suffix
                    Else
                        ' Angka lain biarkan apa adanya (tidak diubah)
                        cell.Value = rawText
                    End If
                End If
            End If
        Next cell
    Next ws
End Sub
