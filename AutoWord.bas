Attribute VB_Name = "AutoWord"
'Automaticaly generate word disp summary info
Sub AutoWordDisp()

    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\AutoReportTemplate.docx", ThisWorkbook.Path & "\AutoReportSource.docx"
    
    wordApp.Documents.Open ThisWorkbook.Path & "\AutoReportSource.docx"

    wordApp.Visible = False

    
    wordApp.ActiveDocument.Bookmarks("dispSummary1").Range.InsertAfter "111"
    

    '测试插入表格可行
    'wordApp.ActiveDocument.Tables.Add wordApp.ActiveDocument.Bookmarks("dispSummary1").Range, NumRows:=14 + 1, NumColumns:=7
    
    'Debug.Print wordApp.ActiveDocument.Variables("tb1").value
    wordApp.ActiveDocument.Variables("tb1").value = "testTb1"
    wordApp.ActiveDocument.Fields.Update
    'wordApp.Selection.Paste    '将剪贴板的内容粘帖到"Select"的部分
    
'        Dim bk
'        For Each bk In wordApp.ActiveDocument.Bookmarks
'            bk.Delete
'        Next
'
'        Dim oFld
'        For Each oFld In ActiveDocument.Fields
'            If oFld.Type = wdFieldDocVariable Then
'                 oFld.Update
'            End If
'        Next oFld
    wordApp.Documents.Save
    
    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\AutoReportResult.docx"
    wordApp.Documents.Close
    wordApp.Quit
    
    Set wordApp = Nothing
    
    
End Sub
