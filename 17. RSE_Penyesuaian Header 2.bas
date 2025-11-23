Sub FormatTabelDanTable_AllSheets()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        For r = 1 To lastRow
            For c = 1 To lastCol
                Select Case Trim(ws.Cells(r, c).Value)
                    Case "Tabel"
                        With ws.Cells(r, c)
                            .HorizontalAlignment = xlLeft
                            .VerticalAlignment = xlBottom
                            .Font.Bold = True
                            .Font.Italic = False
                        End With
                    Case "Table"
                        With ws.Cells(r, c)
                            .HorizontalAlignment = xlLeft
                            .VerticalAlignment = xlTop
                            .Font.Italic = True
                            .Font.Bold = False
                        End With
                End Select
            Next c
        Next r
    Next ws
End Sub
