Sub GabungDokumenKayakPDF()
    Dim dlg As FileDialog
    Dim docUtama As Document
    Dim docSumber As Document
    Dim i As Integer
    
    Set docUtama = ActiveDocument
    
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    With dlg
        .AllowMultiSelect = True
        .Title = "Pilih dokumen yang akan digabung"
        .Filters.Clear
        .Filters.Add "Word Documents", "*.docx; *.doc"
        
        If .Show = -1 Then
            For i = 1 To .SelectedItems.Count
                Set docSumber = Documents.Open(.SelectedItems(i), ReadOnly:=True)
                
                docSumber.Content.Copy
                docUtama.Range(docUtama.Content.End - 1).PasteAndFormat wdFormatOriginalFormatting
                docUtama.Range(docUtama.Content.End - 1).InsertBreak Type:=wdPageBreak
                
                docSumber.Close SaveChanges:=False
            Next i
        End If
    End With
    
    MsgBox "Selesai gabung dokumen kayak PDF!", vbInformation
End Sub
