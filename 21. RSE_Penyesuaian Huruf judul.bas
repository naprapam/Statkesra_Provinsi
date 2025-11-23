Sub FormatPartialText_AllSheets()
    Dim ws As Worksheet
    Dim r As Long, c As Long
    Dim pos As Long
    Dim val As String
    
    For Each ws In ThisWorkbook.Worksheets
        For r = 1 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                val = ws.Cells(r, c).Value
                If InStr(1, val, "Continued Table", vbTextCompare) > 0 Then
                    pos = InStr(1, val, "Continued Table", vbTextCompare)
                    ws.Cells(r, c).Characters(Start:=pos, Length:=Len("Continued Table")).Font.Italic = True
                    ws.Cells(r, c).Characters(Start:=pos, Length:=Len("Continued Table")).Font.Bold = False
                End If
            Next c
        Next r
    Next ws
End Sub
