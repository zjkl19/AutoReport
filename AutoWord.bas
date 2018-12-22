Attribute VB_Name = "AutoWord"
'例子代码
'测试可行，在书签后插入文字
'wordApp.ActiveDocument.Bookmarks("dispSummary1").Range.InsertAfter "111"
'测试插入表格可行
'wordApp.ActiveDocument.Tables.Add wordApp.ActiveDocument.Bookmarks("dispSummary1").Range, NumRows:=14 + 1, NumColumns:=7

'各个统计参数索引
Public Const MaxElasticDeform_Index As Integer = 1
Public Const MinCheckoutCoff_Index As Integer = 2
Public Const MaxCheckoutCoff_Index As Integer = 3
Public Const MinRefRemainDeform_Index As Integer = 4
Public Const MaxRefRemainDeform_Index As Integer = 5

'自动生成计算书
Sub AutoCalcReport()
    Dim wordApp As Word.Application
    Dim doc As Word.Document
        
    Dim i, j As Integer
        
    Dim dispRawTbTitle(1 To MAX_NWC) As String
    Dim dispRawTbTitleVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        dispRawTbTitleVar(i) = Replace("dispRawTbTitle" & GlobalWC(i), " ", "")
    Next
    
    
    Dim dispRawTblBookmarks(1 To MAX_NWC)
    For i = 1 To MAX_NWC
        dispRawTblBookmarks(i) = Replace("dispRawTb" & GlobalWC(i), " ", "")
    Next
        
    Dim dispTbTitle(1 To MAX_NWC) As String
    Dim dispTbTitleVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        dispTbTitleVar(i) = Replace("dispTbTitle" & GlobalWC(i), " ", "")
    Next
    
    Dim dispTblBookmarks(1 To MAX_NWC)
    For i = 1 To MAX_NWC
        dispTblBookmarks(i) = Replace("dispTb" & GlobalWC(i), " ", "")
    Next
    
    
    Dim strainRawTableTitle(1 To MAX_NWC) As String
    Dim strainRawTableTitleVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        strainRawTableTitleVar(i) = Replace("strainRawTableTitle" & StrainGlobalWC(i), " ", "")
    Next
    
    
    Dim strainRawTableBookmarks(1 To MAX_NWC)
    For i = 1 To MAX_NWC
        strainRawTableBookmarks(i) = Replace("strainRawTable" & StrainGlobalWC(i), " ", "")
    Next
    
    Dim strainTableTitle(1 To MAX_NWC) As String
    Dim strainTableTitleVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        strainTableTitleVar(i) = Replace("strainTableTitle" & StrainGlobalWC(i), " ", "")
    Next
    
    
    Dim strainTableBookmarks(1 To MAX_NWC)
    For i = 1 To MAX_NWC
        strainTableBookmarks(i) = Replace("strainTable" & StrainGlobalWC(i), " ", "")
    Next
    
    Dim tbl As Table
    
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\AutoCalcReportTemplate.docx", ThisWorkbook.Path & "\AutoCalcReportSource.docx"
    
    wordApp.Documents.Open ThisWorkbook.Path & "\AutoCalcReportSource.docx"

    wordApp.Visible = False

    For i = 1 To nWCs    'i定位工况
        
        '插入表格标题
        dispRawTbTitle(i) = "表x-x 工况" & CStr(nPN(GlobalWC(i) - 1)) & "挠度原始数据处理表"
        wordApp.ActiveDocument.Variables(dispRawTbTitleVar(i)).value = dispRawTbTitle(i)
        
        '插入表格
        
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(dispRawTblBookmarks(i)).Range, NumRows:=DispUbound(i) + 1, NumColumns:=7)    'NumRows+1表示表头
        tbl.Cell(1, 1).Range.InsertAfter "测点号"
        tbl.Cell(1, 2).Range.InsertAfter "初始读数"
        tbl.Cell(1, 3).Range.InsertAfter "满载"
        tbl.Cell(1, 4).Range.InsertAfter "退载"
        tbl.Cell(1, 5).Range.InsertAfter "总挠度"
        tbl.Cell(1, 6).Range.InsertAfter "弹性挠度"
        tbl.Cell(1, 7).Range.InsertAfter "残余变形"
        

        For j = 1 To DispUbound(i)    'j定位测点
            tbl.Cell(1 + j, 1).Range.InsertAfter CStr(NodeName(i, j))
            tbl.Cell(1 + j, 2).Range.InsertAfter Format(InitDisp(i, j), "Fixed")
            tbl.Cell(1 + j, 3).Range.InsertAfter Format(FullLoadDisp(i, j), "Fixed")
            tbl.Cell(1 + j, 4).Range.InsertAfter Format(UnLoadDisp(i, j), "Fixed")
            tbl.Cell(1 + j, 5).Range.InsertAfter Format(TotalDisp(i, j), "Fixed")
            tbl.Cell(1 + j, 6).Range.InsertAfter Format(ElasticDisp(i, j), "Fixed")
            tbl.Cell(1 + j, 7).Range.InsertAfter Format(RemainDisp(i, j), "Fixed")
        Next
        
        SetTableBorder tbl
    Next i
    
    For i = 1 To nWCs    'i定位工况
        
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
  
        '插入表格标题
        strainRawTableTitle(i) = "表x-x 工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & "应变原始数据处理表"
        wordApp.ActiveDocument.Variables(strainRawTableTitleVar(i)).value = strainRawTableTitle(i)
        
        '插入表格
        
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(strainRawTableBookmarks(i)).Range, NumRows:=StrainUbound(i) + 1, NumColumns:=14)    'NumRows+1表示表头
        tbl.Cell(1, 1).Range.InsertAfter "测点号"
        tbl.Cell(1, 2).Range.InsertAfter "模数R"
        tbl.Cell(1, 3).Range.InsertAfter "温度T"
        tbl.Cell(1, 4).Range.InsertAfter "模数R"
        tbl.Cell(1, 5).Range.InsertAfter "温度T"
        tbl.Cell(1, 6).Range.InsertAfter "模数R"
        tbl.Cell(1, 7).Range.InsertAfter "温度T"
        tbl.Cell(1, 8).Range.InsertAfter "ΔR"
        tbl.Cell(1, 9).Range.InsertAfter "ΔT"
        tbl.Cell(1, 10).Range.InsertAfter "ΔR"
        tbl.Cell(1, 11).Range.InsertAfter "ΔT"
        tbl.Cell(1, 12).Range.InsertAfter "总应变"
        tbl.Cell(1, 13).Range.InsertAfter "弹性应变"
        tbl.Cell(1, 14).Range.InsertAfter "残余应变"
        
        For j = 1 To StrainUbound(i)    'j定位测点
            tbl.Cell(1 + j, 1).Range.InsertAfter CStr((StrainNodeName(i, j)))
            tbl.Cell(1 + j, 2).Range.InsertAfter Format(InitStrainR0(i, j), "#0.0")
            tbl.Cell(1 + j, 3).Range.InsertAfter Format(InitStrainT0(i, j), "#0.0")
            tbl.Cell(1 + j, 4).Range.InsertAfter Format(FullLoadStrainR0(i, j), "#0.0")
            tbl.Cell(1 + j, 5).Range.InsertAfter Format(FullLoadStrainT0(i, j), "#0.0")
            tbl.Cell(1 + j, 6).Range.InsertAfter Format(UnLoadStrainR0(i, j), "#0.0")
            tbl.Cell(1 + j, 7).Range.InsertAfter Format(UnLoadStrainT0(i, j), "#0.0")
            tbl.Cell(1 + j, 8).Range.InsertAfter Format(FullLoadStrainR(i, j), "#0.0")
            tbl.Cell(1 + j, 9).Range.InsertAfter Format(FullLoadStrainT(i, j), "#0.0")
            tbl.Cell(1 + j, 10).Range.InsertAfter Format(UnLoadStrainR(i, j), "#0.0")
            tbl.Cell(1 + j, 11).Range.InsertAfter Format(UnLoadStrainT(i, j), "#0.0")
            tbl.Cell(1 + j, 12).Range.InsertAfter Format(TotalStrain(i, j), "#0.0")
            tbl.Cell(1 + j, 13).Range.InsertAfter Format(ElasticStrain(i, j), "#0.0")
            tbl.Cell(1 + j, 14).Range.InsertAfter Format(RemainStrain(i, j), "#0.0")
        Next
        
        SetTableBorder tbl

    Next

    For i = 1 To StrainNWCs    'i定位工况
  
  
        '插入表格标题
        strainTableTitle(i) = "表x-x 工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & "应变检测结果汇总表"
        wordApp.ActiveDocument.Variables(strainTableTitleVar(i)).value = strainTableTitle(i)
        
        '插入表格
        
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(strainTableBookmarks(i)).Range, NumRows:=StrainUbound(i) + 1, NumColumns:=8)    'NumRows+1表示表头
        tbl.Cell(1, 1).Range.InsertAfter "测点号"
        tbl.Cell(1, 2).Range.InsertAfter "总应变"
        tbl.Cell(1, 3).Range.InsertAfter "弹性应变"
        tbl.Cell(1, 4).Range.InsertAfter "残余应变"
        tbl.Cell(1, 5).Range.InsertAfter "满载应力理论值（MPa）"
        tbl.Cell(1, 6).Range.InsertAfter "满载理论值(με)"
        tbl.Cell(1, 7).Range.InsertAfter "校验系数"
        tbl.Cell(1, 8).Range.InsertAfter "相对残余应变(%)"
        
        For j = 1 To StrainUbound(i)    'j定位测点
            tbl.Cell(1 + j, 1).Range.InsertAfter CStr((StrainNodeName(i, j)))
            tbl.Cell(1 + j, 2).Range.InsertAfter Format(TotalStrain(i, j), "#0.0")
            tbl.Cell(1 + j, 3).Range.InsertAfter Format(ElasticStrain(i, j), "#0.0")
            tbl.Cell(1 + j, 4).Range.InsertAfter Format(RemainStrain(i, j), "#0.0")
            tbl.Cell(1 + j, 5).Range.InsertAfter Format(TheoryStress(i, j), "#0.0")
            tbl.Cell(1 + j, 6).Range.InsertAfter Format(TheoryStrain(i, j), "#0.0")
            tbl.Cell(1 + j, 7).Range.InsertAfter Format(StrainCheckoutCoff(i, j), "Fixed")
            tbl.Cell(1 + j, 8).Range.InsertAfter Format(RefRemainStrain(i, j), "Percent")
        Next
        
        SetTableBorder tbl

    Next
    
    
    wordApp.ActiveDocument.Fields.Update    '更新域

    wordApp.Documents.Save
    
    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\AutoCalcReportResult.docx"
    
    Set wordApp = Nothing
    Set tbl = Nothing
End Sub

'请先计算，再生成Word报告
Public Sub AutoReport()
    'On Error GoTo CloseWord
    
    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    
    Dim i, j As Integer
    
    Dim dispResult(1 To MAX_NWC) As String    '和word中响应DocVariable对应
    Dim dispResultVar(1 To MAX_NWC) As String    '和word中响应DocVariable对应（常量数组）
    For i = 1 To MAX_NWC
        dispResultVar(i) = Replace("dispResult" & GlobalWC(i), " ", "")
    Next
    
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
    
    Dim strainResult(1 To MAX_NWC) As String
    Dim strainResultVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        strainResultVar(i) = Replace("strainResult" & StrainGlobalWC(i), " ", "")
    Next
    
    Dim strainSummary(1 To MAX_NWC) As String
    Dim strainSummaryVar(1 To MAX_NWC) As String
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
        '插入结果描述
        dispResult(i) = "(" & Str(i) & ")在工况" & CStr(nPN(GlobalWC(i) - 1)) & "荷载作用下，主梁最大实测弹性挠度值为" & Format(StatPara(i, MaxElasticDeform_Index), "Fixed") & "mm，" _
        & "实测控制截面的挠度值均小于理论值，校验系数在" & Format(StatPara(i, MinCheckoutCoff_Index), "Fixed") & "～" & Format(StatPara(i, MaxCheckoutCoff_Index), "Fixed") & "之间；" _
        & "相对残余变形在" & Format(StatPara(i, MinRefRemainDeform_Index), "Percent") & "～" & Format(StatPara(i, MaxRefRemainDeform_Index), "Percent") & "之间。"
        
        wordApp.ActiveDocument.Variables(dispResultVar(i)).value = dispResult(i)
    
        '插入概述
        dispSummary(i) = "工况" & CStr(nPN(GlobalWC(i) - 1)) & "测试截面测点挠度检测结果详见表x-x、图x-x。检测结果表明，所测主梁的挠度校验系数在" & Format(StatPara(i, MinCheckoutCoff_Index), "Fixed") & "～" & Format(StatPara(i, MaxCheckoutCoff_Index), "Fixed") & "之间，" _
        & "满足《公路桥梁承载能力检测评定规程》中规定的校验系数小于1.0的要求。所测主梁的最大相对残余变形为" & Format(StatPara(i, MaxRefRemainDeform_Index), "Percent") & "，" _
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
        '插入结果描述
        strainResult(i) = "(" & Str(i) & ")在工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & "荷载作用下，所测主梁最大弹性应变为" & Format(StrainStatPara(i, MaxElasticDeform_Index), "Fixed") & "με，" _
        & "实测控制截面的混凝土应变值均小于理论值，校验系数在" & Format(StrainStatPara(i, MinCheckoutCoff_Index), "Fixed") & "～" & Format(StrainStatPara(i, MaxCheckoutCoff_Index), "Fixed") & "之间；" _
        & "相对残余应变在" & Format(StrainStatPara(i, MinRefRemainDeform_Index), "Percent") & "～" & Format(StrainStatPara(i, MaxRefRemainDeform_Index), "Percent") & "之间。"
        
        wordApp.ActiveDocument.Variables(strainResultVar(i)).value = strainResult(i)
        '插入概述
        strainSummary(i) = "工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & "测试截面测点应变检测结果详见表x-x、图x-x。检测结果表明，所测主梁的应变校验系数在" _
        & Format(StrainStatPara(i, MinCheckoutCoff_Index), "Fixed") & "～" & Format(StrainStatPara(i, MaxCheckoutCoff_Index), "Fixed") & "之间，" _
        & "满足《公路桥梁承载能力检测评定规程》规定的校验系数小于1.0的要求。所测构件的最大相对残余应变为" & Format(StrainStatPara(i, MaxRefRemainDeform_Index), "Percent") _
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
            tbl.Cell(1 + j, 2).Range.InsertAfter Format(INTTotalStrain(i, j), "Fixed")
            tbl.Cell(1 + j, 3).Range.InsertAfter Format(INTElasticStrain(i, j), "Fixed")
            tbl.Cell(1 + j, 4).Range.InsertAfter Format(INTRemainStrain(i, j), "Fixed")
            tbl.Cell(1 + j, 5).Range.InsertAfter Format(TheoryStrain(i, j), "Fixed")
            tbl.Cell(1 + j, 6).Range.InsertAfter Format(INTDivStrainCheckoutCoff(i, j), "Fixed")
            tbl.Cell(1 + j, 7).Range.InsertAfter Format(INTDivRefRemainStrain(i, j), "Percent")
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
    Set tbl = Nothing
    
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
