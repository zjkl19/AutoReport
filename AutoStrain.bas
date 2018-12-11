Attribute VB_Name = "AutoStrain"
Option Explicit
Dim FirstRow As Integer    '起始数据行数
Dim TotalDispCol As Integer    '总变形所在列
Dim DeltaCol As Integer   '增量所在列
Dim RemainDispCol As Integer    '残余变形所在列
Dim ElasticCol As Integer    '弹性变形所在列
Dim CheckoutCoffCol As Integer    '校验系数所在列
Dim RefRemainDisp As Integer    '相对残余变形所在列

Sub AutoDisp_Click()

    Dim rowCurr As Integer    '行指针
    FirstRow = 3
    
    TotalDispCol = 5
    DeltaCol = 7
    RemainDispCol = 8
    ElasticCol = 9
    CheckoutCoffCol = 11
    RefRemainDisp = 12
    
    rowCurr = FirstRow
    While Cells(rowCurr, 1) <> ""
        Cells(rowCurr, TotalDispCol) = Cells(rowCurr, TotalDispCol - 1) - Cells(rowCurr, TotalDispCol - 2)
        Cells(rowCurr, DeltaCol) = Cells(rowCurr, DeltaCol - 1) - Cells(rowCurr, DeltaCol - 3)
        
        '算法：卸载与满载读数差值>=0，取卸载与满载读数差值，否则取0
        Cells(rowCurr, RemainDispCol) = IIf(Cells(rowCurr, RemainDispCol - 2) - Cells(rowCurr, RemainDispCol - 5) >= 0, Cells(rowCurr, RemainDispCol - 2) - Cells(rowCurr, RemainDispCol - 5), 0)
        
        Cells(rowCurr, ElasticCol) = Cells(rowCurr, ElasticCol - 4) - Cells(rowCurr, ElasticCol - 1)
        Cells(rowCurr, CheckoutCoffCol) = Cells(rowCurr, CheckoutCoffCol - 2) / Cells(rowCurr, CheckoutCoffCol - 1)
        Cells(rowCurr, RefRemainDisp) = Cells(rowCurr, RefRemainDisp - 4) / Cells(rowCurr, RefRemainDisp - 7)
        
        rowCurr = rowCurr + 1
    Wend
 
End Sub


Sub test()
    Dim wCondition(100)   '例：wCondition(0)='一',wCondition(1)='二'
    Dim wPointer As Integer
    
    Dim i As Integer
    Dim rowCurr As Integer    '行指针
    
    For i = 0 To 99
        wCondition(i) = ""
    Next
    
    FirstRow = 3
    rowCurr = FirstRow
    
    i = 0
    '判断工况数（不重复的数字算不同的工况）
    While Cells(rowCurr, 1) <> ""
        If i = 0 Then
            wCondition(i) = Cells(rowCurr, 1)
        ElseIf Cells(rowCurr, 1) <> wCondition(i) Then    '有新工况
            i = i + 1
            wCondition(i) = Cells(rowCurr, 1)
        End If
        
        rowCurr = rowCurr + 1
    Wend
    
    Dim nWK    '工况数
    i = 0
    While wCondition(i) <> ""
        i = i + 1
    Wend
    nWK = i

    
    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    
    Set wordApp = CreateObject("Word.Application")
    'Set doc = wordApp.Documents.Add
    'Set doc = wordapp.Documents.Open("AutoReport.docx")
    'Set r = doc.Range(Start:=0, End:=0)
    
    'r.InsertAfter "test"
    Dim tbl As Table
    Set tbl = wordApp.ActiveDocument.Tables.Add(Range:=wordApp.Documents.Add.Range(Start:=0, End:=0), NumRows:=14 + 1, NumColumns:=7)
    tbl.Cell(1, 1).Range.InsertAfter "测点号"
    tbl.Cell(1, 2).Range.InsertAfter "总变形"
    tbl.Cell(1, 3).Range.InsertAfter "弹性变形"
    tbl.Cell(1, 4).Range.InsertAfter "残余变形"
    tbl.Cell(1, 5).Range.InsertAfter "满载理论值(mm)"
    tbl.Cell(1, 6).Range.InsertAfter "校验系数"
    tbl.Cell(1, 7).Range.InsertAfter "相对残余变形(%)"
    
    TotalDispCol = 5
    DeltaCol = 7
    RemainDispCol = 8
    ElasticCol = 9
    CheckoutCoffCol = 11
    RefRemainDisp = 12
    
    For i = 1 To 14
        tbl.Cell(1 + i, 1).Range.InsertAfter Format(Cells(2 + i, 2), "Fixed")
        tbl.Cell(1 + i, 2).Range.InsertAfter Format(Cells(2 + i, TotalDispCol), "Fixed")
        tbl.Cell(1 + i, 3).Range.InsertAfter Format(Cells(2 + i, ElasticCol), "Fixed")
        tbl.Cell(1 + i, 4).Range.InsertAfter Format(Cells(2 + i, RemainDispCol), "Fixed")
        tbl.Cell(1 + i, 5).Range.InsertAfter Format(Cells(2 + i, 10), "Fixed")
        tbl.Cell(1 + i, 6).Range.InsertAfter Format(Cells(2 + i, CheckoutCoffCol), "Fixed")
        tbl.Cell(1 + i, 7).Range.InsertAfter Format(Cells(2 + i, RefRemainDisp), "Percent")
    Next
    
    'Set t = wordApp.ActiveDocument.Tables(1)    '表格序号从1开始算？
    'tbl.Cell(1, 1).Range.InsertAfter "第一个单元格" '第一个单元格中写入字符"第一个单元格"
    'tbl.Cell(tbl.Rows.Count, tbl.Columns.Count).Range.InsertAfter "最后一个单元格" '第二个单元格中写入字符"最后一个单元格"
        
    With tbl
        With .Borders
            .InsideLineStyle = wdLineStyleSingle
            .OutsideLineStyle = wdLineStyleSingle
        End With
    End With

    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\AutoReport.docx"
    wordApp.Documents.Close
    wordApp.Quit
    
    Set wordApp = Nothing
    
     
End Sub


