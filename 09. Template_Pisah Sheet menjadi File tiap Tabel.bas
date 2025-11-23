Sub SplitSheetsByBab()
    Dim ws As Worksheet
    Dim wbNew As Workbook
    Dim dict As Object
    Dim prefix As String
    Dim wb As Workbook
    
    Set wb = ThisWorkbook
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Loop semua sheet
    For Each ws In wb.Worksheets
        ' Ambil angka depan sebelum titik
        If InStr(ws.Name, ".") > 0 Then
            prefix = Left(ws.Name, InStr(ws.Name, ".") - 1)
        Else
            prefix = ws.Name
        End If
        
        ' Kalau prefix belum ada di dictionary, buat workbook baru
        If Not dict.Exists(prefix) Then
            Set wbNew = Workbooks.Add
            dict.Add prefix, wbNew
        Else
            Set wbNew = dict(prefix)
        End If
        
        ' Copy sheet ke workbook sesuai prefix
        ws.Copy After:=wbNew.Sheets(wbNew.Sheets.Count)
        
        ' Hapus sheet kosong default kalau masih ada
        On Error Resume Next
        Application.DisplayAlerts = False
        wbNew.Sheets("Sheet1").Delete
        Application.DisplayAlerts = True
        On Error GoTo 0
    Next ws
    
    ' Simpan semua workbook hasil
    Dim key As Variant
    For Each key In dict.Keys
        dict(key).SaveAs wb.Path & "\Bab " & key & ".xlsx", FileFormat:=51
        dict(key).Close SaveChanges:=False
    Next key
    
    MsgBox "Selesai! File dipisah berdasarkan Bab."
End Sub
