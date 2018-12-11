Attribute VB_Name = "AutoDisp"
Option Explicit
Const First_Row As Integer = 8     '起始数据行数
Const WC_Col As Integer = 1     '工况所在列数
Const MAX_NWC As Integer = 10     '最大工况数

Dim TotalDispCol As Integer    '总变形所在列
Dim DeltaCol As Integer   '增量所在列
Dim RemainDispCol As Integer    '残余变形所在列
Dim ElasticCol As Integer    '弹性变形所在列
Dim CheckoutCoffCol As Integer    '校验系数所在列
Dim RefRemainDispCol As Integer    '相对残余变形所在列

Dim nWCs As Integer    '工况数
Dim TotalDisp(1 To MAX_NWC, 1 To 100)   'TotalDisp(i,j)表示第i个工况，第j个测点总变形
Dim Delta
Dim RemainDisp
Dim ElasticDisp
Dim CheckoutCoff
Dim RefRemainDisp
Dim DispUbound(1 To MAX_NWC) As Integer

Dim t

'''初始化全局变量
Public Sub InitVar()
    nWCs = Cells(1, 2)
    Dim i As Integer
    For i = 1 To nWCs
        DispUbound(i) = Cells(2, 2 * i)
    Next
    
    TotalDispCol = 5
    DeltaCol = 7
    RemainDispCol = 8
    ElasticCol = 9
    CheckoutCoffCol = 11
    RefRemainDispCol = 12
End Sub


Private Sub AutoDisp_Click()

    InitVar
    
    Dim rowCurr As Integer    '行指针
    

    
    rowCurr = First_Row
    While Cells(rowCurr, 1) <> ""
        Cells(rowCurr, TotalDispCol) = Cells(rowCurr, TotalDispCol - 1) - Cells(rowCurr, TotalDispCol - 2)
        Cells(rowCurr, DeltaCol) = Cells(rowCurr, DeltaCol - 1) - Cells(rowCurr, DeltaCol - 3)
        
        '算法：卸载与满载读数差值>=0，取卸载与满载读数差值，否则取0
        Cells(rowCurr, RemainDispCol) = IIf(Cells(rowCurr, RemainDispCol - 2) - Cells(rowCurr, RemainDispCol - 5) >= 0, Cells(rowCurr, RemainDispCol - 2) - Cells(rowCurr, RemainDispCol - 5), 0)
        
        Cells(rowCurr, ElasticCol) = Cells(rowCurr, ElasticCol - 4) - Cells(rowCurr, ElasticCol - 1)
        Cells(rowCurr, CheckoutCoffCol) = Cells(rowCurr, CheckoutCoffCol - 2) / Cells(rowCurr, CheckoutCoffCol - 1)
        Cells(rowCurr, RefRemainDispCol) = Cells(rowCurr, RefRemainDisp - 4) / Cells(rowCurr, RefRemainDispCol - 7)
        
        rowCurr = rowCurr + 1
    Wend

 
End Sub
'''计算最小/最大校验系数，最大相对残余变形等参数
Sub CalcParas()
    Dim nWCs As Integer    '工况数
    Dim nPs(10) As Integer    '各个工况测点数
    Dim nPN     '各个工况对应中文名称
    nPN = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十")
    Dim i, j As Integer
    nWCs = Cells(1, 2)
    For i = 0 To nWCs - 1
        nPs(i) = Cells(2, 2 * (i + 1))
    Next
    
    Dim rowCurr As Integer    '行指针
    

    
    For i = 0 To nWCs - 1    '遍历工况
        rowCurr = 3    '每个工况开始，行指针重新开始
        For j = 0 To 2 '
            If j = 0 Then
                Cells(rowCurr, 2 * (i + 1)) = 0
            ElseIf j = 1 Then
                Cells(rowCurr, 2 * (i + 1)) = 0
            ElseIf j = 2 Then
                Cells(rowCurr, 2 * (i + 1)) = 0
            End If
  
            rowCurr = rowCurr + 1
        Next
    Next
End Sub
'''生成各个工况对应行
Private Sub GenerateRows_Click()
    Dim nWCs As Integer    '工况数
    Dim nPs(10) As Integer    '各个工况测点数
    Dim nPN     '各个工况对应中文名称
    nPN = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十")
    Dim i, j As Integer
    nWCs = Cells(1, 2)
    For i = 0 To nWCs - 1
        nPs(i) = Cells(2, 2 * (i + 1))
        'Debug.Print nPs(i)
    Next
    'Debug.Print nWCs
    
    
    Dim rowCurr As Integer    '行指针
    rowCurr = First_Row
    
    For i = 0 To nWCs - 1    '遍历工况
        For j = 1 To nPs(i)    '遍历各个工况的测点
            Cells(rowCurr, WC_Col) = nPN(i)
            rowCurr = rowCurr + 1
        Next
    Next
    
 
End Sub
Sub test()
    Dim wCondition(100)   '例：wCondition(0)='一',wCondition(1)='二'
    Dim wPointer As Integer
    
    Dim i As Integer
    Dim rowCurr As Integer    '行指针
    
    For i = 0 To 99
        wCondition(i) = ""
    Next
    
    First_Row = 3
    rowCurr = First_Row
    
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
    RefRemainDispCol = 12
    
    For i = 1 To 14
        tbl.Cell(1 + i, 1).Range.InsertAfter Format(Cells(2 + i, 2), "Fixed")
        tbl.Cell(1 + i, 2).Range.InsertAfter Format(Cells(2 + i, TotalDispCol), "Fixed")
        tbl.Cell(1 + i, 3).Range.InsertAfter Format(Cells(2 + i, ElasticCol), "Fixed")
        tbl.Cell(1 + i, 4).Range.InsertAfter Format(Cells(2 + i, RemainDispCol), "Fixed")
        tbl.Cell(1 + i, 5).Range.InsertAfter Format(Cells(2 + i, 10), "Fixed")
        tbl.Cell(1 + i, 6).Range.InsertAfter Format(Cells(2 + i, CheckoutCoffCol), "Fixed")
        tbl.Cell(1 + i, 7).Range.InsertAfter Format(Cells(2 + i, RefRemainDispCol), "Percent")
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


