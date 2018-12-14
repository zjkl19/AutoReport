Attribute VB_Name = "AutoWord"
'例子代码
'测试可行，在书签后插入文字
'wordApp.ActiveDocument.Bookmarks("dispSummary1").Range.InsertAfter "111"
'测试插入表格可行
'wordApp.ActiveDocument.Tables.Add wordApp.ActiveDocument.Bookmarks("dispSummary1").Range, NumRows:=14 + 1, NumColumns:=7


'请先计算，再生成Word报告
Sub AutoWordDisp()
    'On Error GoTo CloseWord
    
    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    
    Dim i, j As Integer
    Dim dispSummary(1 To MAX_NWC) As String    '和word中响应DocVariable对应
    Dim dispSummaryVar(1 To MAX_NWC) As String    '和word中响应DocVariable对应（常量数组）
    For i = 1 To MAX_NWC
        dispSummaryVar(i) = Replace("dispSummary" & GlobalWC(i), " ", "")
    Next
    
    Dim dispTbTitle(1 To MAX_NWC) As String
    Dim dispTbTitleVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        dispTbTitleVar(i) = Replace("dispTbTitle" & GlobalWC(i), " ", "")
    Next
    
    Dim strainSummary(1 To MAX_NWC) As String    '和word中响应DocVariable对应
    Dim strainSummaryVar(1 To MAX_NWC) As String    '和word中响应DocVariable对应
    For i = 1 To MAX_NWC
        strainSummaryVar(i) = Replace("strainSummary" & StrainGlobalWC(i), " ", "")
    Next
    
    Dim strainTbTitle(1 To MAX_NWC) As String
    Dim strainTbTitleVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        strainTbTitleVar(i) = Replace("strainTbTitle" & StrainGlobalWC(i), " ", "")
    Next
    
    Dim dispTblBookmarks(1 To MAX_NWC)
    For i = 1 To MAX_NWC
        dispTblBookmarks(i) = Replace("dispTable" & GlobalWC(i), " ", "")
    Next
    
    Dim strainTblBookmarks(1 To MAX_NWC)
    For i = 1 To MAX_NWC
        strainTblBookmarks(i) = Replace("strainTable" & StrainGlobalWC(i), " ", "")
    Next

    Dim tbl As Table

    
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\AutoReportTemplate.docx", ThisWorkbook.Path & "\AutoReportSource.docx"
    
    wordApp.Documents.Open ThisWorkbook.Path & "\AutoReportSource.docx"

    wordApp.Visible = False

    For i = 1 To nWCs    'i定位工况
    
        '插入概述
        dispSummary(i) = "工况" & CStr(nPN(GlobalWC(i) - 1)) & "测试截面测点挠度检测结果详见表x-x、图x-x。检测结果表明，所测主梁的挠度校验系数在" & Format(StatPara(i, 1), "Fixed") & "～" & Format(StatPara(i, 2), "Fixed") & "之间，" _
        & "满足《公路桥梁承载能力检测评定规程》中规定的校验系数小于1.0的要求。所测主梁的最大相对残余变形为" & Format(StatPara(i, 3), "Percent") & "，" _
        & "满足《公路桥梁承载能力检测评定规程》中规定的残余变形限值要求(限值20%)，恢复状况良好。"
               
        wordApp.ActiveDocument.Variables(dispSummaryVar(i)).value = dispSummary(i)
        
        '插入表格标题
        dispTbTitle(i) = "表x-x 工况" & CStr(nPN(GlobalWC(i) - 1)) & "挠度检测结果汇总表"
        wordApp.ActiveDocument.Variables(dispTbTitleVar(i)).value = dispTbTitle(i)
        
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
        
        SetTableBorder tbl
    Next i
    
    For i = 1 To StrainNWCs    'i定位工况
    
        '插入概述
        strainSummary(i) = "工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & "测试截面测点应变检测结果详见表x-x、图x-x。检测结果表明，所测主梁的应变校验系数在" _
        & Format(StrainStatPara(i, 1), "Fixed") & "～" & Format(StrainStatPara(i, 2), "Fixed") & "之间，" _
        & "满足《公路桥梁承载能力检测评定规程》规定的校验系数小于1.0的要求。所测构件的最大相对残余应变为" & Format(StrainStatPara(i, 3), "Percent") _
        & "，满足《公路桥梁承载能力检测评定规程》中规定的残余应变限值要求(限值20%)，恢复状况良好。"
        
        wordApp.ActiveDocument.Variables(strainSummaryVar(i)).value = strainSummary(i)
        
        '插入表格标题
        strainTbTitle(i) = "表x-x 工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & "应变检测结果汇总表"
        wordApp.ActiveDocument.Variables(strainTbTitleVar(i)).value = strainTbTitle(i)
        
        '插入表格
        
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(strainTblBookmarks(i)).Range, NumRows:=StrainUbound(i) + 1, NumColumns:=7)    'NumRows+1表示表头
        tbl.Cell(1, 1).Range.InsertAfter "测点号"
        tbl.Cell(1, 2).Range.InsertAfter "总应变"
        tbl.Cell(1, 3).Range.InsertAfter "弹性应变"
        tbl.Cell(1, 4).Range.InsertAfter "残余应变"
        tbl.Cell(1, 5).Range.InsertAfter "满载理论值(με)"
        tbl.Cell(1, 6).Range.InsertAfter "校验系数"
        tbl.Cell(1, 7).Range.InsertAfter "相对残余应变(%)"
        
        For j = 1 To StrainUbound(i)    'j定位测点
            tbl.Cell(1 + j, 1).Range.InsertAfter Format(StrainNodeName(i, j), "Fixed")
            tbl.Cell(1 + j, 2).Range.InsertAfter Format(TotalStrain(i, j), "Fixed")
            tbl.Cell(1 + j, 3).Range.InsertAfter Format(ElasticStrain(i, j), "Fixed")
            tbl.Cell(1 + j, 4).Range.InsertAfter Format(RemainStrain(i, j), "Fixed")
            tbl.Cell(1 + j, 5).Range.InsertAfter Format(TheoryStrain(i, j), "Fixed")
            tbl.Cell(1 + j, 6).Range.InsertAfter Format(StrainCheckoutCoff(i, j), "Fixed")
            tbl.Cell(1 + j, 7).Range.InsertAfter Format(RefRemainStrain(i, j), "Percent")
        Next
        
        SetTableBorder tbl

    Next
    
    wordApp.ActiveDocument.Fields.Update    '更新域

    wordApp.Documents.Save
    
    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\AutoReportResult.docx"
CloseWord:
    wordApp.Documents.Close
    wordApp.Quit
    
    Set wordApp = Nothing
    
    
End Sub
'设置table的边界线
Private Sub SetTableBorder(ByRef tbl As Table)
    With tbl
        With .Borders
            .InsideLineStyle = wdLineStyleSingle
            .OutsideLineStyle = wdLineStyleSingle
        End With
    End With
End Sub
