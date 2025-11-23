Sub HapusBaris1dan3_AllSheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ' hapus baris 1
        ws.Rows(1).Delete
        ' hapus baris 3 (yang tadinya baris 3, bergeser jadi baris 3 setelah baris 1 dihapus)
        ws.Rows(3).Delete
    Next ws
End Sub
