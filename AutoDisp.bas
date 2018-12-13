Attribute VB_Name = "AutoDisp"
Option Explicit
Const First_Row As Integer = 8     '起始数据行数
Const WC_Col As Integer = 1     '工况所在列数
Const MAX_NWC As Integer = 10     '最大工况数

Const Node_Name_Col As Integer = 2  '测点编号所在列
Const TheoryDisp_Col As Integer = 10  '理论位移所在列
Dim TotalDispCol As Integer    '总变形所在列
Dim DeltaCol As Integer   '增量所在列
Dim RemainDispCol As Integer    '残余变形所在列
Dim ElasticCol As Integer    '弹性变形所在列
Dim CheckoutCoffCol As Integer    '校验系数所在列
Dim RefRemainDispCol As Integer    '相对残余变形所在列

Dim nWCs As Integer    '工况数
Dim nPN    '各个工况对应中文名称
'nPN = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十")

Dim TotalDisp(1 To MAX_NWC, 1 To 100)   'TotalDisp(i,j)表示第i个工况，第j个测点总变形
Dim NodeName(1 To MAX_NWC, 1 To 100) As String  '各个工况测点名称
Dim Delta(1 To MAX_NWC, 1 To 100)
Dim RemainDisp(1 To MAX_NWC, 1 To 100)
Dim ElasticDisp(1 To MAX_NWC, 1 To 100)
Dim TheoryDisp(1 To MAX_NWC, 1 To 100)
Dim CheckoutCoff(1 To MAX_NWC, 1 To 100)
Dim RefRemainDisp(1 To MAX_NWC, 1 To 100)
Dim DispUbound(1 To MAX_NWC) As Integer    '每个工况上界（下界为1）

Dim StatPara(1 To MAX_NWC, 1 To 3)  '统计参数,最小校验系数，最大校验系数，最大相对残余应变
'StatPara(i,1~3)分别表示第i个工况最小校验系数，最大校验系数，最大相对残余应变

Dim t

'''初始化全局变量
Public Sub InitVar()

    nPN = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十")
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
    
    Dim i, j As Integer
    rowCurr = First_Row
    For i = 1 To nWCs
        For j = 1 To DispUbound(i)
        
            NodeName(i, j) = Cells(rowCurr, Node_Name_Col)
            TheoryDisp(i, j) = Cells(rowCurr, TheoryDisp_Col)
            
            TotalDisp(i, j) = Cells(rowCurr, TotalDispCol - 1) - Cells(rowCurr, TotalDispCol - 2)    '总变形
            Cells(rowCurr, TotalDispCol) = TotalDisp(i, j)
            
            Cells(rowCurr, DeltaCol) = Cells(rowCurr, DeltaCol - 1) - Cells(rowCurr, DeltaCol - 3)   '增量
            '增量不存储
            
            '算法：卸载与满载读数差值>=0，取卸载与满载读数差值，否则取0
            RemainDisp(i, j) = IIf(Cells(rowCurr, RemainDispCol - 2) - Cells(rowCurr, RemainDispCol - 5) >= 0, Cells(rowCurr, RemainDispCol - 2) - Cells(rowCurr, RemainDispCol - 5), 0)
            Cells(rowCurr, RemainDispCol) = RemainDisp(i, j)    '残余变形
            
            ElasticDisp(i, j) = Cells(rowCurr, ElasticCol - 4) - Cells(rowCurr, ElasticCol - 1)
            Cells(rowCurr, ElasticCol) = ElasticDisp(i, j)    '弹性变形
             
            CheckoutCoff(i, j) = Cells(rowCurr, CheckoutCoffCol - 2) / Cells(rowCurr, CheckoutCoffCol - 1)
            Cells(rowCurr, CheckoutCoffCol) = CheckoutCoff(i, j)    '校验系数
             
            RefRemainDisp(i, j) = Cells(rowCurr, RefRemainDispCol - 4) / Cells(rowCurr, RefRemainDispCol - 7)
            Cells(rowCurr, RefRemainDispCol) = RefRemainDisp(i, j)    '相对残余变形
            
            rowCurr = rowCurr + 1
        Next
    Next
    
    '计算各个工况最小/大校验系数，最大相对残余变形
    For i = 1 To nWCs
        StatPara(i, 1) = CheckoutCoff(i, 1): StatPara(i, 2) = CheckoutCoff(i, 1): StatPara(i, 3) = RefRemainDisp(i, 1)
        For j = 1 To DispUbound(i)
            If (CheckoutCoff(i, j) < StatPara(i, 1)) Then
                StatPara(i, 1) = CheckoutCoff(i, j)
            End If
            If (CheckoutCoff(i, j) > StatPara(i, 2)) Then
                StatPara(i, 2) = CheckoutCoff(i, j)
            End If
            If (RefRemainDisp(i, j) > StatPara(i, 3)) Then
                StatPara(i, 3) = RefRemainDisp(i, j)
            End If
        Next
        
        '数据写入Excel
        Cells(3, 2 * i) = Format(StatPara(i, 1), "Fixed"): Cells(4, 2 * i) = Format(StatPara(i, 2), "Fixed"): Cells(5, 2 * i) = Format(StatPara(i, 3), "Percent")
    Next
'    While Cells(rowCurr, 1) <> ""
'        Cells(rowCurr, TotalDispCol) = Cells(rowCurr, TotalDispCol - 1) - Cells(rowCurr, TotalDispCol - 2)
'        Cells(rowCurr, DeltaCol) = Cells(rowCurr, DeltaCol - 1) - Cells(rowCurr, DeltaCol - 3)
'
'        算法：卸载与满载读数差值>=0，取卸载与满载读数差值，否则取0
'        Cells(rowCurr, RemainDispCol) = IIf(Cells(rowCurr, RemainDispCol - 2) - Cells(rowCurr, RemainDispCol - 5) >= 0, Cells(rowCurr, RemainDispCol - 2) - Cells(rowCurr, RemainDispCol - 5), 0)
'
'        Cells(rowCurr, ElasticCol) = Cells(rowCurr, ElasticCol - 4) - Cells(rowCurr, ElasticCol - 1)
'        Cells(rowCurr, CheckoutCoffCol) = Cells(rowCurr, CheckoutCoffCol - 2) / Cells(rowCurr, CheckoutCoffCol - 1)
'        Cells(rowCurr, RefRemainDispCol) = Cells(rowCurr, RefRemainDisp - 4) / Cells(rowCurr, RefRemainDispCol - 7)
'
'        rowCurr = rowCurr + 1
'    Wend

 
End Sub
'''计算最小/最大校验系数，最大相对残余变形等参数
'Sub CalcParas()
'    Dim nWCs As Integer    '工况数
'    Dim nPs(10) As Integer    '各个工况测点数
'
'    nPN = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十")
'    Dim i, j As Integer
'    nWCs = Cells(1, 2)
'    For i = 0 To nWCs - 1
'        nPs(i) = Cells(2, 2 * (i + 1))
'    Next
'
'    Dim rowCurr As Integer    '行指针
'
'
'
'    For i = 0 To nWCs - 1    '遍历工况
'        rowCurr = 3    '每个工况开始，行指针重新开始
'        For j = 0 To 2 '
'            If j = 0 Then
'                Cells(rowCurr, 2 * (i + 1)) = 0
'            ElseIf j = 1 Then
'                Cells(rowCurr, 2 * (i + 1)) = 0
'            ElseIf j = 2 Then
'                Cells(rowCurr, 2 * (i + 1)) = 0
'            End If
'
'            rowCurr = rowCurr + 1
'        Next
'    Next
'End Sub
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

Sub AutoWordDisp()
    'On Error GoTo CloseWord
    AutoDisp_Click
    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    
    Dim i, j As Integer
    Dim dispSummary(1 To 10) As String    '和word中响应DocVariable对应
    Dim dispSummaryVar    '和word中响应DocVariable对应
    dispSummaryVar = Array("dispSummary1", "dispSummary2", "dispSummary3", "dispSummary4", "dispSummary5", "dispSummary6", "dispSummary7", "dispSummary8", "dispSummary9", "dispSummary10")
    Dim dispTblBookmarks(1 To 10)
    For i = 1 To 10
        dispTblBookmarks(i) = Replace("dispTable" & Str(i), " ", "")
    Next

        Dim tbl As Table

    
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\AutoReportTemplate.docx", ThisWorkbook.Path & "\AutoReportSource.docx"
    
    wordApp.Documents.Open ThisWorkbook.Path & "\AutoReportSource.docx"

    wordApp.Visible = False

    For i = 1 To nWCs    'i定位工况
    
        '插入概述
        dispSummary(i) = "工况" & CStr(nPN(i - 1)) & "测试截面测点挠度检测结果详见表x-x、图x-x。检测结果表明，所测主梁的挠度校验系数在" & Format(StatPara(i, 1), "Fixed") & "～" & Format(StatPara(i, 2), "Fixed") & "之间，" _
        & "满足《公路桥梁承载能力检测评定规程》中规定的校验系数小于1.0的要求。所测主梁的最大相对残余变形为" & Format(StatPara(i, 3), "Percent") & "，" _
        & "满足《公路桥梁承载能力检测评定规程》中规定的残余变形限值要求(限值20%)，恢复状况良好。"
        
        wordApp.ActiveDocument.Variables(dispSummaryVar(i - 1)).value = dispSummary(i)
        wordApp.ActiveDocument.Fields.Update
        
        '插入表格
        
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(dispTblBookmarks(i)).Range, NumRows:=DispUbound(i) + 1, NumColumns:=7)    'NumRows+1表示表头
        tbl.Cell(1, 1).Range.InsertAfter "测点号"
        tbl.Cell(1, 2).Range.InsertAfter "总变形"
        tbl.Cell(1, 3).Range.InsertAfter "弹性变形"
        tbl.Cell(1, 4).Range.InsertAfter "残余变形"
        tbl.Cell(1, 5).Range.InsertAfter "满载理论值(mm)"
        tbl.Cell(1, 6).Range.InsertAfter "校验系数"
        tbl.Cell(1, 7).Range.InsertAfter "相对残余变形(%)"
        

        For j = 1 To DispUbound(i)    'j定位测点
            tbl.Cell(1 + j, 1).Range.InsertAfter Format(NodeName(i, j), "Fixed")
            tbl.Cell(1 + j, 2).Range.InsertAfter Format(TotalDisp(i, j), "Fixed")
            tbl.Cell(1 + j, 3).Range.InsertAfter Format(ElasticDisp(i, j), "Fixed")
            tbl.Cell(1 + j, 4).Range.InsertAfter Format(RemainDisp(i, j), "Fixed")
            tbl.Cell(1 + j, 5).Range.InsertAfter Format(TheoryDisp(i, j), "Fixed")
            tbl.Cell(1 + j, 6).Range.InsertAfter Format(CheckoutCoff(i, j), "Fixed")
            tbl.Cell(1 + j, 7).Range.InsertAfter Format(RefRemainDisp(i, j), "Percent")
        Next
        
        With tbl
            With .Borders
                .InsideLineStyle = wdLineStyleSingle
                .OutsideLineStyle = wdLineStyleSingle
            End With
        End With
    Next i
    
    'Dim TotalDisp(1 To MAX_NWC, 1 To 100)   'TotalDisp(i,j)表示第i个工况，第j个测点总变形
    
    'Dim NodeName(1 To MAX_NWC, 1 To 100) As String  '各个工况测点名称
    'Dim Delta(1 To MAX_NWC, 1 To 100)
    'Dim RemainDisp(1 To MAX_NWC, 1 To 100)
    'Dim ElasticDisp(1 To MAX_NWC, 1 To 100)
    'Dim CheckoutCoff(1 To MAX_NWC, 1 To 100)
    'Dim RefRemainDisp(1 To MAX_NWC, 1 To 100)
    'Dim DispUbound(1 To MAX_NWC) As Integer    '每个工况上界（下界为1）
    '
    'Dim StatPara(1 To MAX_NWC, 1 To 3)  '统计参数,最小校验系数，最大校验系数，最大相对残余应变
    ''StatPara(i,1~3)分别表示第i个工况最小校验系数，最大校验系数，最大相对残余应变
    '
    'Dim t
    

    


    '测试可行，在书签后插入文字
    'wordApp.ActiveDocument.Bookmarks("dispSummary1").Range.InsertAfter "111"
    

    '测试插入表格可行
    'wordApp.ActiveDocument.Tables.Add wordApp.ActiveDocument.Bookmarks("dispSummary1").Range, NumRows:=14 + 1, NumColumns:=7
    
    'Debug.Print wordApp.ActiveDocument.Variables("tb1").value
    
    
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
CloseWord:
    wordApp.Documents.Close
    wordApp.Quit
    
    Set wordApp = Nothing
    
    
End Sub


Sub GenerateWordReport()
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


