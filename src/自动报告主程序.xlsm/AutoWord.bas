Attribute VB_Name = "AutoWord"
Option Explicit
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

Public Const AutoReportFileName As String = "自动生成的报告.docx" '荷载试验报告
Public Const AutoCalcReportFileName As String = "自动生成的计算书.docx" '荷载试验计算书

'测试自动生成的模板（报告模板、计算书模板）
Sub TestGenBookmarkAndDocVar()
    Dim resultFlag As Boolean
    resultFlag = False
    
    Dim templateFileName As String
    'templateFileName = "空自动报告模板.docx"            '"报告模板.docx"
    templateFileName = "桥梁静动载试验报告模板.docx"
    Dim fileName As String
    fileName = "空自动报告.docx"
    
    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    Dim r1 As Word.Range
    
    Dim i, j As Integer
        
    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    FileCopy ThisWorkbook.Path & "\" & templateFileName, ThisWorkbook.Path & "\S.docx"
    
    wordApp.Documents.Open ThisWorkbook.Path & "\S.docx"

    wordApp.Visible = False
    
    With wordApp.ActiveDocument.Application
        .Selection.TypeParagraph
        Set r = .Selection.Range
        '        'r.MoveEnd , -1

        r.Text = "DOCVARIABLE dispResult1 \* MERGEFORMAT"
        r.Fields.Add r, wdFieldEmpty, , False
        r.Font.name = "楷体_GB2312"                '参考代码：'Selection.Font.Name = "楷体_GB2312"
        r.ParagraphFormat.Alignment = wdAlignParagraphLeft
        '        .Selection.Text = "DOCVARIABLE dispResult1 \* MERGEFORMAT"
        '        .Selection.Range.Fields.Add .Selection.Range, wdFieldEmpty, , False
        '        .Selection.Font.Name = "楷体_GB2312"
        '
        '        r.MoveEnd , -1
        '
        .Selection.EndKey
        .Selection.TypeParagraph
        '.Selection.TypeText Text:="word"
        Set r = .Selection.Range
        r.Text = "DOCVARIABLE dispResult2 \* MERGEFORMAT"
        r.Fields.Add r, wdFieldEmpty, , False
        r.Font.name = "楷体_GB2312"                '参考代码：'Selection.Font.Name = "楷体_GB2312"
        r.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.EndKey
        .Selection.TypeParagraph
        '        'Set r = Nothing
        '
        '        .Selection.Text = "DOCVARIABLE dispResult2 \* MERGEFORMAT"
        '        .Selection.Range.Fields.Add .Selection.Range, wdFieldEmpty, , False
        '        .Selection.Font.Name = "楷体_GB2312"
        
        '        .Selection.MoveEnd

        'Set r1 = .Selection.Range
        'r.MoveEnd , -1
        '         r.Text = "DOCVARIABLE dispResult2 \* MERGEFORMAT"
        '         r.Fields.Add r, wdFieldEmpty, , False
        '         r.Font.Name = "楷体_GB2312"    '参考代码：'Selection.Font.Name = "楷体_GB2312"

        '参考代码
        '.Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Name:="2"
        '.Selection.MoveDown Unit:=wdLine, Count:=5
        '.Selection.MoveRight Unit:=wdCharacter, Count:=9
        '.Selection.TypeText Text:="word"    '参考代码
    End With
    Set r = Nothing

    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    With wordApp.ActiveDocument.Bookmarks
        .Add Range:=wordApp.ActiveDocument.Application.Selection.Range, name:="lbt"
        .DefaultSorting = wdSortByName
        .ShowHidden = False
    End With

    '     With wordApp.ActiveDocument.Bookmarks
    '        .Add Range:=Selection.Range, Name:="lbt1"
    '        .DefaultSorting = wdSortByLocation
    '        .ShowHidden = False
    '    End With
    '     wordApp.ActiveDocument.Application.Selection
    '    wordApp.ActiveDocument.Selection.MoveDown 4, 1
    'wordApp.ActiveDocument.Fields.Update    '更新域

    wordApp.Documents.Save
    
    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\" & fileName
    resultFlag = True
    
CloseWord:

    wordApp.Documents.Close
    wordApp.Quit
    Set wordApp = Nothing
    Set r = Nothing
    
    If resultFlag = True Then
        MsgBox "导出完成！"
    Else
        MsgBox "导出失败！"
    End If
End Sub

'测试自动生成的模板（报告模板、计算书模板）
Sub testGenTemplate()
    Dim templateFileName As String
    templateFileName = "空自动报告模板.docx"            '"报告模板.docx"
    
    Dim fileName As String
    fileName = "空自动报告.docx"
    
    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim i, j As Integer
        
        
    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    FileCopy ThisWorkbook.Path & "\" & templateFileName, ThisWorkbook.Path & "\S.docx"
    
    wordApp.Documents.Open ThisWorkbook.Path & "\S.docx"

    wordApp.Visible = False
    
    With wordApp.ActiveDocument
        .Selection.GoTo what:=wdGoToPage, Which:=wdGoToNext, name:="2"
        .Selection.MoveDown Unit:=wdLine, Count:=5
        .Selection.MoveRight Unit:=wdCharacter, Count:=9
        .Selection.TypeText Text:="word"
    End With
 
    wordApp.ActiveDocument.Fields.Update         '更新域

    wordApp.Documents.Save
    
    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\" & fileName
    resultFlag = True
    
CloseWord:

    wordApp.Documents.Close
    wordApp.Quit
    Set wordApp = Nothing
    Set tbl = Nothing
    If resultFlag = True Then
        MsgBox "导出完成！"
    Else
        MsgBox "导出失败！"
    End If
End Sub

'设置报告"正文"样式格式
Public Sub SetReportMainBodyFormat(ByRef wordApp As Word.Application)
    With wordApp.Application.ActiveDocument.Styles.Item("正文")
        With .Font
            .NameAscii = "Times New Roman"
            .NameOther = "Times New Roman"
            .NameFarEast = "楷体_GB2312"
            .Size = 12                           '小四
        End With
    End With
End Sub

'设置报告"题注"样式格式
Public Sub SetReportCaptionsFormat(ByRef wordApp As Word.Application)
    With wordApp.Application.ActiveDocument.Styles.Item("题注")
        With .Font
            .NameAscii = "Times New Roman"
            .NameOther = "Times New Roman"
            .NameFarEast = "楷体_GB2312"
            .Size = 12                           '小四
        End With
    End With
End Sub

'设置自动报告的题注和交叉引用
Public Sub SetAutoReportCaptions(ByRef wordApp As Word.Application)
    Dim i As Integer

    '挠度
    For i = 1 To NWCs
        With wordApp.Application
            With .Selection.Find
                .Text = dispTbTitle(i)
                .Forward = True
                .Wrap = wdFindContinue
                .MatchCase = False
                .MatchByte = True
                .MatchWildcards = False
                .MatchWholeWord = False
                .MatchFuzzy = False
                .Replacement.Text = ""
            End With
            .Selection.Find.Execute Replace:=wdReplaceNone
            .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
            .Selection.TypeText " "
            .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
            With .CaptionLabels("表")
                .NumberStyle = wdCaptionNumberStyleArabic
                .IncludeChapterNumber = -1
                .ChapterStyleLevel = 1
                .Separator = wdSeparatorHyphen
            End With
            .Selection.InsertCaption Label:="表", Title:="", TitleAutoText:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=False
            
        End With
    Next
    
    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then '注：只能使用xlOn，不能使用True（尽管真值为1）
        For i = 1 To NWCs
            With wordApp.Application
                With .Selection.Find
                    .Text = dispChartTitle(i)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchCase = False
                    .MatchByte = True
                    .MatchWildcards = False
                    .MatchWholeWord = False
                    .MatchFuzzy = False
                    .Replacement.Text = ""
                End With
                .Selection.Find.Execute Replace:=wdReplaceNone
                .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
                .Selection.TypeText " "
                .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
                With .CaptionLabels("图")
                    .NumberStyle = wdCaptionNumberStyleArabic
                    .IncludeChapterNumber = -1
                    .ChapterStyleLevel = 1
                    .Separator = wdSeparatorHyphen
                End With
                .Selection.InsertCaption Label:="图", Title:="", TitleAutoText:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=False
            End With
        Next
    End If
    
    '应变
    For i = 1 To NWCs
        With wordApp.Application
            With .Selection.Find
                .Text = strainTbTitle(i)
                .Forward = True
                .Wrap = wdFindContinue
                .MatchCase = False
                .MatchByte = True
                .MatchWildcards = False
                .MatchWholeWord = False
                .MatchFuzzy = False
                .Replacement.Text = ""
            End With
            .Selection.Find.Execute Replace:=wdReplaceNone
            .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
            .Selection.TypeText " "
            .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
            With .CaptionLabels("表")
                .NumberStyle = wdCaptionNumberStyleArabic
                .IncludeChapterNumber = -1
                .ChapterStyleLevel = 1
                .Separator = wdSeparatorHyphen
            End With
            .Selection.InsertCaption Label:="表", Title:="", TitleAutoText:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=False
        End With
    Next
    
    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then '注：只能使用xlOn，不能使用True（尽管真值为1）
        For i = 1 To NWCs
            With wordApp.Application
                With .Selection.Find
                    .Text = strainChartTitle(i)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchCase = False
                    .MatchByte = True
                    .MatchWildcards = False
                    .MatchWholeWord = False
                    .MatchFuzzy = False
                    .Replacement.Text = ""
                End With
                .Selection.Find.Execute Replace:=wdReplaceNone
                .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
                .Selection.TypeText " "
                .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
                With .CaptionLabels("图")
                    .NumberStyle = wdCaptionNumberStyleArabic
                    .IncludeChapterNumber = -1
                    .ChapterStyleLevel = 1
                    .Separator = wdSeparatorHyphen
                End With
                .Selection.InsertCaption Label:="图", Title:="", TitleAutoText:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=False
            End With
        Next
    End If
    

End Sub

'设置图或表的题注
'searchText:查找字符
'captionName:题注标签名
Public Sub AddCaptions(ByRef wordApp As Word.Application, ByVal searchText As String, ByVal captionName As String)
        With wordApp.Application
            With .Selection.Find
                .Text = searchText
                .Forward = True
                .Wrap = wdFindContinue
                .MatchCase = False
                .MatchByte = True
                .MatchWildcards = False
                .MatchWholeWord = False
                .MatchFuzzy = False
                .Replacement.Text = ""
            End With
            .Selection.Find.Execute Replace:=wdReplaceNone
            .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
            .Selection.TypeText " "
            .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
            With .CaptionLabels(captionName)
                .NumberStyle = wdCaptionNumberStyleArabic
                .IncludeChapterNumber = -1
                .ChapterStyleLevel = 1
                .Separator = wdSeparatorHyphen
            End With
            .Selection.InsertCaption Label:=captionName, Title:="", TitleAutoText:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=False
            
        End With
End Sub

'设置计算书自动报告的题注和交叉引用
Public Sub SetAutoCalcReportCaptions(ByRef wordApp As Word.Application)
    Dim i As Integer
    '挠度
    For i = 1 To NWCs
        AddCaptions wordApp, dispRawTbTitle(i), "表"
        AddCaptions wordApp, dispTbTitle(i), "表"
    Next
    
    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox_Calc") = xlOn Then '注：只能使用xlOn，不能使用True（尽管真值为1）
        For i = 1 To NWCs
            AddCaptions wordApp, dispChartTitle(i), "图"
        Next
    End If
    
    For i = 1 To NWCs
        AddCaptions wordApp, dispTheoryShapeTitle(i), "图"
    Next
    
    '应变
    For i = 1 To strainNWCs
        AddCaptions wordApp, strainRawTbTitle(i), "表"
        AddCaptions wordApp, strainTbTitle(i), "表"
    Next
    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox_Calc") = xlOn Then '注：只能使用xlOn，不能使用True（尽管真值为1）
        For i = 1 To strainNWCs
               AddCaptions wordApp, strainChartTitle(i), "图"
        Next
    End If
    
    For i = 1 To strainNWCs
        AddCaptions wordApp, strainTheoryShapeTitle(i), "图"
    Next
    
End Sub

Public Sub SetAutoReportCrossReferences(ByRef wordApp As Word.Application, Optional tableOffset As Integer = 0, Optional chartOffset As Integer = 0)
    Dim i As Integer
    Dim rs As ReportService
    Set rs = New ReportService
    
    For i = 1 To NWCs
        With wordApp.Application
            '插入交叉引用
            With .Selection.Find
                .Text = dispTbCrossRef(i)
                .Forward = True
                .Wrap = wdFindContinue
                .MatchCase = False
                .MatchByte = True
                .MatchWildcards = False
                .MatchWholeWord = False
                .MatchFuzzy = False
                .Replacement.Text = ""
            End With
            .Selection.Find.Execute Replace:=wdReplaceNone
            .Selection.TypeBackspace
            .Selection.InsertCrossReference ReferenceType:="表", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=rs.GetDispCrossReferenceItem(i, GlobalWC, StrainGlobalWC, strainNWCs, tableOffset), _
        InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
        End With
        If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then '注：只能使用xlOn，不能使用True（尽管真值为1）
            With wordApp.Application
                With .Selection.Find
                    .Text = dispGraphCrossRef(i)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchCase = False
                    .MatchByte = True
                    .MatchWildcards = False
                    .MatchWholeWord = False
                    .MatchFuzzy = False
                    .Replacement.Text = ""
                End With
                .Selection.Find.Execute Replace:=wdReplaceNone
                .Selection.TypeBackspace
                .Selection.InsertCrossReference ReferenceType:="图", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=rs.GetDispCrossReferenceItem(i, GlobalWC, StrainGlobalWC, strainNWCs, chartOffset), _
            InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
            End With
        End If
    Next
    
    For i = 1 To strainNWCs
        With wordApp.Application
            '插入交叉引用
            With .Selection.Find
                .Text = strainTbCrossRef(i)
                .Forward = True
                .Wrap = wdFindContinue
                .MatchCase = False
                .MatchByte = True
                .MatchWildcards = False
                .MatchWholeWord = False
                .MatchFuzzy = False
                .Replacement.Text = ""
            End With
            .Selection.Find.Execute Replace:=wdReplaceNone
            .Selection.TypeBackspace
            .Selection.InsertCrossReference ReferenceType:="表", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=rs.GetStrainCrossReferenceItem(i, GlobalWC, NWCs, StrainGlobalWC, tableOffset), _
        InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
        End With
        If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then '注：只能使用xlOn，不能使用True（尽管真值为1）
            With wordApp.Application
                With .Selection.Find
                    .Text = strainGraphCrossRef(i)
                    .Forward = True
                    .Wrap = wdFindContinue
                    .MatchCase = False
                    .MatchByte = True
                    .MatchWildcards = False
                    .MatchWholeWord = False
                    .MatchFuzzy = False
                    .Replacement.Text = ""
                End With
                .Selection.Find.Execute Replace:=wdReplaceNone
                .Selection.TypeBackspace
                .Selection.InsertCrossReference ReferenceType:="图", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=rs.GetStrainCrossReferenceItem(i, GlobalWC, NWCs, StrainGlobalWC, chartOffset), _
            InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
            End With
        End If
    Next
    
    Set rs = Nothing
End Sub

'生成自动报告模板
Public Sub GenAutoReportTemplate()
        
    Dim ax As ArrayService                       '原as变量名与关键字as冲突，改为as
    Set ax = New ArrayService
    
    Dim resultFlag As Boolean
    resultFlag = False
    
    Dim templateFileName As String
    templateFileName = "桥梁静动载试验报告模板.docx"           '"报告模板.docx"
    
    Dim fileName As String
    fileName = "自动报告模板.docx"
    
    Const ReportStartBookmarkName As String = "ReportStart"
    Const DispResultStartBookmarkName  As String = "DispResultStart"
    Const StrainResultStartBookmarkName  As String = "StrainResultStart"
    
    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    Dim r1 As Word.Range
    
    Dim i As Integer
    Dim j As Integer
        
    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    FileCopy ThisWorkbook.Path & "\" & templateFileName, ThisWorkbook.Path & "\S.docx"
    
    wordApp.Documents.Open ThisWorkbook.Path & "\S.docx"

    wordApp.Visible = False
    
    wordApp.ActiveDocument.Application.Selection.GoTo what:=wdGoToBookmark, name:=DispResultStartBookmarkName
    wordApp.ActiveDocument.Application.Selection.TypeParagraph
    For i = 1 To NWCs                            'i定位工况
        With wordApp.ActiveDocument.Application
            Set r = .Selection.Range
            '        'r.MoveEnd , -1
    
            r.Text = "dispResult" & CStr(i)
            'r.Font.Name = "Times New Roman"            '参考代码：'Selection.Font.Name = "楷体_GB2312"
            r.ParagraphFormat.Alignment = wdAlignParagraphLeft
            r.ParagraphFormat.CharacterUnitFirstLineIndent = 2
            r.Fields.Add r, wdFieldDocVariable, r.Text, True
            '        .Selection.Text = "DOCVARIABLE dispResult1 \* MERGEFORMAT"
            '        .Selection.Range.Fields.Add .Selection.Range, wdFieldEmpty, , False
            '        .Selection.Font.Name = "楷体_GB2312"
            '
            '        r.MoveEnd , -1
            '
            .Selection.EndKey
            .Selection.TypeParagraph
            '.Selection.TypeText Text:="word"
    
        End With
    Next i
    
    wordApp.ActiveDocument.Application.Selection.TypeParagraph
    wordApp.ActiveDocument.Application.Selection.GoTo what:=wdGoToBookmark, name:=StrainResultStartBookmarkName
    For i = 1 To strainNWCs
        With wordApp.ActiveDocument.Application
            Set r = .Selection.Range

            r.Text = "strainResult" & CStr(i)
            'r.Font.Name = "楷体_GB2312"
            r.ParagraphFormat.Alignment = wdAlignParagraphLeft
            r.ParagraphFormat.CharacterUnitFirstLineIndent = 2
            r.Fields.Add r, wdFieldDocVariable, r.Text, True
            
            .Selection.EndKey
            .Selection.TypeParagraph

        End With
    Next i
    
    wordApp.ActiveDocument.Application.Selection.TypeParagraph
    
    wordApp.ActiveDocument.Application.Selection.GoTo what:=wdGoToBookmark, name:=ReportStartBookmarkName
    '计算报告工况总数
    Dim totalNWC As Integer                      '总工况数
    totalNWC = ax.ArrayMax(GlobalWC)
    If totalNWC < ax.ArrayMax(StrainGlobalWC) Then '比较挠度、应变最大工况
        totalNWC = ax.ArrayMax(StrainGlobalWC)
    End If
    'Debug.Print totalNWC
    
    For i = 1 To totalNWC
        With wordApp.ActiveDocument.Application
            '遍历挠度所有工况，若报告工况对应，则写入
            For j = 1 To NWCs
                If GlobalWC(j) = i Then
                    Set r = .Selection.Range
                        
                    r.Text = "dispSummary" & CStr(j)
                    'r.Font.Name = "楷体_GB2312"
                    r.ParagraphFormat.Alignment = wdAlignParagraphLeft
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 2
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                        
                    .Selection.EndKey
                    .Selection.TypeParagraph
                        
                    Set r = .Selection.Range
                    r.Text = "dispTbTitle" & CStr(j)
                    'r.Font.Name = "楷体_GB2312"
                    r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                    
                    .Selection.EndKey
                    .Selection.TypeParagraph
                        
                    '挠度表书签
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    With wordApp.ActiveDocument.Bookmarks
                        .Add Range:=wordApp.ActiveDocument.Application.Selection.Range, name:="dispTable" & Trim(CStr(j))
                        .DefaultSorting = wdSortByName
                        .ShowHidden = False
                    End With
                        
                    .Selection.EndKey
                    .Selection.TypeParagraph
                    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then
                        '挠度图书签
                        wordApp.ActiveDocument.Application.Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
                        wordApp.ActiveDocument.Application.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                        wordApp.ActiveDocument.Application.Selection.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                        wordApp.ActiveDocument.Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        With wordApp.ActiveDocument.Bookmarks
                            .Add Range:=wordApp.ActiveDocument.Application.Selection.Range, name:="dispChart" & Trim(CStr(j))
                            .DefaultSorting = wdSortByName
                            .ShowHidden = False
                        End With
                        .Selection.EndKey
                        .Selection.TypeParagraph
                        
                        '挠度图标题
                        Set r = .Selection.Range
                        r.Text = "dispChartTitle" & CStr(j)
                        'r.Font.Name = "楷体_GB2312"
                        r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                        r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                        r.Fields.Add r, wdFieldDocVariable, r.Text, True
                        
                        .Selection.EndKey
                        .Selection.TypeParagraph
                    End If
                End If
            Next j
            For j = 1 To strainNWCs
                If StrainGlobalWC(j) = i Then
                    Set r = .Selection.Range
            
                    r.Text = "strainSummary" & CStr(j)
                    'r.Font.Name = "楷体_GB2312"
                    r.ParagraphFormat.Alignment = wdAlignParagraphLeft
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 2
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                        
                    .Selection.EndKey
                    .Selection.TypeParagraph
                        
                    Set r = .Selection.Range
                    r.Text = "strainTbTitle" & CStr(j)
                    'r.Font.Name = "楷体_GB2312"
                    r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                         
                    .Selection.EndKey
                    .Selection.TypeParagraph
                    '应变表书签
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    With wordApp.ActiveDocument.Bookmarks
                        .Add Range:=wordApp.ActiveDocument.Application.Selection.Range, name:="strainTable" & Trim(CStr(j))
                        .DefaultSorting = wdSortByName
                        .ShowHidden = False
                    End With
            
                    .Selection.EndKey
                    .Selection.TypeParagraph
                    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then
                        '应变图书签
                        wordApp.ActiveDocument.Application.Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
                        wordApp.ActiveDocument.Application.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                        wordApp.ActiveDocument.Application.Selection.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                        wordApp.ActiveDocument.Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        With wordApp.ActiveDocument.Bookmarks
                            .Add Range:=wordApp.ActiveDocument.Application.Selection.Range, name:="strainChart" & Trim(CStr(j))
                            .DefaultSorting = wdSortByName
                            .ShowHidden = False
                        End With
                        .Selection.EndKey
                        .Selection.TypeParagraph
                        
                        Set r = .Selection.Range
                        r.Text = "strainChartTitle" & CStr(j)
                        'r.Font.Name = "楷体_GB2312"
                        r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                        r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                        r.Fields.Add r, wdFieldDocVariable, r.Text, True
                    End If
                         
                    .Selection.EndKey
                    .Selection.TypeParagraph
                End If
            Next j
                        
        End With
           
    Next
    
    Set r = Nothing
    
    '     wordApp.ActiveDocument.Application.Selection
    '    wordApp.ActiveDocument.Selection.MoveDown 4, 1
    'wordApp.ActiveDocument.Fields.Update    '更新域

    wordApp.Documents.Save
    
    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\" & fileName
    resultFlag = True
    
CloseWord:

    wordApp.Documents.Close
    wordApp.Quit
    Set wordApp = Nothing
    Set r = Nothing
    Set ax = Nothing
    
    'TODO：模板生成失败，则终止过程
    '    If resultFlag = True Then
    '        MsgBox "导出完成！"
    '    Else
    '        MsgBox "导出失败！"
    '    End If
End Sub

'GenAutoCalcReportTemplate过程中设置具体内容
Private Sub SetAutoCalcReportTemplateDetail(ByRef wordApp As Word.Application, ByVal totalNWC As Integer)
    
    Const CalcReportStartBookmarkName As String = "CalcReportStart"
    Dim r As Word.Range
    Dim i As Integer
    Dim j As Integer
    wordApp.Visible = False
    wordApp.ActiveDocument.Application.Selection.GoTo what:=wdGoToBookmark, name:=CalcReportStartBookmarkName
    For i = 1 To totalNWC
        With wordApp.ActiveDocument.Application
            '遍历挠度所有工况，若报告工况对应，则写入
            For j = 1 To NWCs
                If GlobalWC(j) = i Then
                    '挠度原始数据处理表标题
                    Set r = .Selection.Range
                    r.Text = "dispRawTbTitle" & CStr(j)
                    'r.Font.Name = "楷体_GB2312"
                    r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                    .Selection.EndKey
                    .Selection.TypeParagraph
                        
                    '挠度原始数据处理表书签
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
                    With wordApp.ActiveDocument.Bookmarks
                        .Add Range:=wordApp.ActiveDocument.Application.Selection.Range, name:="dispRawTb" & Trim(CStr(j))
                        .DefaultSorting = wdSortByName
                        .ShowHidden = False
                    End With
                    .Selection.EndKey
                    .Selection.TypeParagraph
                    
                    '挠度表标题
                    Set r = .Selection.Range
                    r.Text = "dispTbTitle" & CStr(j)
                    'r.Font.Name = "楷体_GB2312"
                    r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                    .Selection.EndKey
                    .Selection.TypeParagraph
                        
                    '挠度表书签
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
                    With wordApp.ActiveDocument.Bookmarks
                        .Add Range:=wordApp.ActiveDocument.Application.Selection.Range, name:="dispTable" & Trim(CStr(j))
                        .DefaultSorting = wdSortByName
                        .ShowHidden = False
                    End With
                    .Selection.EndKey
                    .Selection.TypeParagraph
                    
                    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox_Calc") = xlOn Then
                        '挠度图书签
                        wordApp.ActiveDocument.Application.Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
                        wordApp.ActiveDocument.Application.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                        wordApp.ActiveDocument.Application.Selection.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                        wordApp.ActiveDocument.Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        With wordApp.ActiveDocument.Bookmarks
                            .Add Range:=wordApp.ActiveDocument.Application.Selection.Range, name:="dispChart" & Trim(CStr(j))
                            .DefaultSorting = wdSortByName
                            .ShowHidden = False
                        End With
                        .Selection.EndKey
                        .Selection.TypeParagraph
                    
                        '挠度图标题
                        Set r = .Selection.Range
                        r.Text = "dispChartTitle" & CStr(j)
                        'r.Font.Name = "楷体_GB2312"
                        r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                        r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                        r.Fields.Add r, wdFieldDocVariable, r.Text, True
                        .Selection.EndKey
                        .Selection.TypeParagraph

                    End If
                     '挠度理论图书签
                     With wordApp.ActiveDocument.Application.Selection.ParagraphFormat
                        .LineSpacingRule = wdLineSpaceSingle
                        .CharacterUnitFirstLineIndent = 0
                        .FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                        .Alignment = wdAlignParagraphCenter
                    End With
                    With wordApp.ActiveDocument.Bookmarks
                        .Add Range:=wordApp.ActiveDocument.Application.Selection.Range, name:="dispTheoryShape" & Trim(CStr(j))
                        .DefaultSorting = wdSortByName
                        .ShowHidden = False
                    End With
                    .Selection.EndKey
                    .Selection.TypeParagraph
                    '挠度理论值标题
                    Set r = .Selection.Range
                    r.Text = "dispTheoryShapeTitle" & CStr(j)
                    r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                    .Selection.EndKey
                    .Selection.TypeParagraph
                End If
            Next j
            
            For j = 1 To strainNWCs
                If StrainGlobalWC(j) = i Then
                    '应变原始数据处理表标题
                    Set r = .Selection.Range
                    r.Text = "strainRawTbTitle" & CStr(j)
                    'r.Font.Name = "楷体_GB2312"
                    r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                    .Selection.EndKey
                    .Selection.TypeParagraph
                    '应变原始数据处理表书签
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
                    With wordApp.ActiveDocument.Bookmarks
                        .Add Range:=wordApp.ActiveDocument.Application.Selection.Range, name:="strainRawTable" & Trim(CStr(j))
                        .DefaultSorting = wdSortByName
                        .ShowHidden = False
                    End With
                    .Selection.EndKey
                    .Selection.TypeParagraph
                    '应变表标题
                    Set r = .Selection.Range
                    r.Text = "strainTbTitle" & CStr(j)
                    'r.Font.Name = "楷体_GB2312"
                    r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                    .Selection.EndKey
                    .Selection.TypeParagraph
                    '应变表书签
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    wordApp.ActiveDocument.Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
                    With wordApp.ActiveDocument.Bookmarks
                        .Add Range:=wordApp.ActiveDocument.Application.Selection.Range, name:="strainTable" & Trim(CStr(j))
                        .DefaultSorting = wdSortByName
                        .ShowHidden = False
                    End With
                    .Selection.EndKey
                    .Selection.TypeParagraph
                    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox_Calc") = xlOn Then
                        '应变图书签
                        wordApp.ActiveDocument.Application.Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
                        wordApp.ActiveDocument.Application.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                        wordApp.ActiveDocument.Application.Selection.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                        wordApp.ActiveDocument.Application.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        With wordApp.ActiveDocument.Bookmarks
                            .Add Range:=wordApp.ActiveDocument.Application.Selection.Range, name:="strainChart" & Trim(CStr(j))
                            .DefaultSorting = wdSortByName
                            .ShowHidden = False
                        End With
                        .Selection.EndKey
                        .Selection.TypeParagraph
                        '应变图标题
                        Set r = .Selection.Range
                        r.Text = "strainChartTitle" & CStr(j)
                        'r.Font.Name = "楷体_GB2312"
                        r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                        r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                        r.Fields.Add r, wdFieldDocVariable, r.Text, True
                         
                        .Selection.EndKey
                        .Selection.TypeParagraph
                    End If
                     '应变理论图书签
                     With wordApp.ActiveDocument.Application.Selection.ParagraphFormat
                        .LineSpacingRule = wdLineSpaceSingle
                        .CharacterUnitFirstLineIndent = 0
                        .FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                        .Alignment = wdAlignParagraphCenter
                    End With
                    With wordApp.ActiveDocument.Bookmarks
                        .Add Range:=wordApp.ActiveDocument.Application.Selection.Range, name:="strainTheoryShape" & Trim(CStr(j))
                        .DefaultSorting = wdSortByName
                        .ShowHidden = False
                    End With
                    .Selection.EndKey
                    .Selection.TypeParagraph
                    '应变理论值标题
                    Set r = .Selection.Range
                    r.Text = "strainTheoryShapeTitle" & CStr(j)
                    r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                    .Selection.EndKey
                    .Selection.TypeParagraph
                End If
            Next j
                        
        End With
           
    Next
    
    Set r = Nothing
End Sub

'生成自动报告模板（）
Public Sub GenAutoCalcReportTemplate()
        
    Dim ax As ArrayService                       '原as变量名与关键字as冲突，改为as
    Set ax = New ArrayService
    
    Dim resultFlag As Boolean
    resultFlag = False
    
    Dim templateFileName As String
    'templateFileName = "空自动计算书模板.docx"           '"报告模板.docx"
    templateFileName = "桥梁静动载试验计算书模板.docx"
    Dim fileName As String
    fileName = "自动计算书模板.docx"
    
    Dim wordApp As Word.Application
    Dim doc As Word.Document

    Dim r1 As Word.Range
        
    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    FileCopy ThisWorkbook.Path & "\" & templateFileName, ThisWorkbook.Path & "\temp.docx"
    
    wordApp.Documents.Open ThisWorkbook.Path & "\temp.docx"
    'wordApp.ActiveDocument.Application.Selection.TypeParagraph
    
    '计算报告工况总数
    Dim totalNWC As Integer                      '总工况数
    totalNWC = ax.ArrayMax(GlobalWC)
    If totalNWC < ax.ArrayMax(StrainGlobalWC) Then '比较挠度、应变最大工况
        totalNWC = ax.ArrayMax(StrainGlobalWC)
    End If
    
    SetAutoCalcReportTemplateDetail wordApp, totalNWC
    ' wordApp.ActiveDocument.Application.Selection
    ' wordApp.ActiveDocument.Selection.MoveDown 4, 1
    'wordApp.ActiveDocument.Fields.Update    '更新域

    wordApp.Documents.Save
    
    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\" & fileName
    resultFlag = True
    
CloseWord:

    wordApp.Documents.Close
    wordApp.Quit
    Set wordApp = Nothing
    Set ax = Nothing
    
    'TODO：模板生成失败，则终止过程
    '    If resultFlag = True Then
    '        MsgBox "导出完成！"
    '    Else
    '        MsgBox "导出失败！"
    '    End If
End Sub
'获取理论图的指针
'算法：如果图片左边夹在边界之间，则存储该指针
'sheetName:标签名
'shapeObjArray():图片数组（数组需预留足够大小）
'leftBoundry:左边界
'rightBoundry:右边界
'返回值：图片数量
Public Function GetTheoryGraph(ByVal sheetName As String, ByRef shapeObjArray() As Shape, ByVal leftBoundry As Integer, ByVal rightBoundry As Integer)
Dim i As Integer
i = 0
Dim s As Shape
For Each s In Sheets(sheetName).Shapes
    If s.Left > leftBoundry And s.Left < rightBoundry Then
        i = i + 1
        Set shapeObjArray(i) = s
    End If
Next
ShapeArraySort shapeObjArray, i
GetTheoryGraph = i
End Function

'Sort a Shape Array by the top of the array [SINGLE DIMENSION]
Private Sub ShapeArraySort(ByRef shapeObjArray() As Shape, ByVal arrayCounts As Integer)
    '从小到大排列
    'On Error Resume Next
    Dim outerIndex As Integer
    Dim innerIndex As Integer
    Dim Temp As Shape
    For outerIndex = LBound(shapeObjArray) To arrayCounts
        For innerIndex = outerIndex + 1 To arrayCounts
            
            If shapeObjArray(innerIndex).Top < shapeObjArray(outerIndex).Top Then
                Set Temp = shapeObjArray(innerIndex)
                Set shapeObjArray(innerIndex) = shapeObjArray(outerIndex)
                Set shapeObjArray(outerIndex) = Temp
            End If
            
        Next innerIndex
    Next outerIndex

End Sub


'自动生成计算书
Public Sub AutoCalcReport()
     
    GenAutoCalcReportTemplate
    Dim templateFileName As String
    templateFileName = "自动计算书模板.docx"            '"报告模板.docx"
    
    Dim calcFileName As String
    calcFileName = "自动生成的计算书.docx"

    Dim wordApp As Word.Application
    Dim doc As Word.Document
        
    Dim i, j As Integer
    
    Dim resultFlag As Boolean                    'True表示导出成功
    resultFlag = False
    
    Dim dispChartTitleVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        dispChartTitleVar(i) = Replace("dispChartTitle" & Str(i), " ", "")
    Next
       
    Dim strainChartTitleVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        strainChartTitleVar(i) = Replace("strainChartTitle" & CStr(i), " ", "")
    Next
    
    Dim dispRawTbTitleVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        dispRawTbTitleVar(i) = Replace("dispRawTbTitle" & CStr(i), " ", "")
    Next
    
    Dim dispRawTblBookmarks(1 To MAX_NWC)
    For i = 1 To MAX_NWC
        dispRawTblBookmarks(i) = Replace("dispRawTb" & CStr(i), " ", "")
    Next
        
    Dim dispTbTitleVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        dispTbTitleVar(i) = Replace("dispTbTitle" & CStr(i), " ", "")
    Next
    
    Dim dispTblBookmarks(1 To MAX_NWC)
    For i = 1 To MAX_NWC
        dispTblBookmarks(i) = Replace("dispTable" & CStr(i), " ", "")
    Next
    
    Dim strainRawTableTitleVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        strainRawTableTitleVar(i) = Replace("strainRawTbTitle" & CStr(i), " ", "")
    Next
    
    Dim strainRawTableBookmarks(1 To MAX_NWC)
    For i = 1 To MAX_NWC
        strainRawTableBookmarks(i) = Replace("strainRawTable" & CStr(i), " ", "")
    Next
    
    Dim strainTbTitleVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        strainTbTitleVar(i) = Replace("strainTbTitle" & CStr(i), " ", "")
    Next
    
    Dim strainTableBookmarks(1 To MAX_NWC)
    For i = 1 To MAX_NWC
        strainTableBookmarks(i) = Replace("strainTable" & CStr(i), " ", "")
    Next
    
    '计算截图区边界
    Dim DispTheoryLeftBound As Integer     '截图左边界
    Dim DispTheoryRightBound As Integer      '截图右边界
    DispTheoryLeftBound = 0    '默认表格宽度下为1200左右
    For i = 1 To 23
        DispTheoryLeftBound = DispTheoryLeftBound + Sheets(DispSheetName).Cells(3, i).Width
    Next
    DispTheoryRightBound = 0
    For i = 1 To 28    '默认表格宽度下为1500左右
        DispTheoryRightBound = DispTheoryRightBound + Sheets(DispSheetName).Cells(3, i).Width
    Next
    
    Dim StrainTheoryLeftBound As Integer     '截图左边界
    Dim StrainTheoryRightBound As Integer      '截图右边界
    StrainTheoryLeftBound = 0    '默认表格宽度下为1800左右
    For i = 1 To 33
        StrainTheoryLeftBound = StrainTheoryLeftBound + Sheets(StrainSheetName).Cells(3, i).Width
    Next
    StrainTheoryRightBound = 0
    For i = 1 To 38    '默认表格宽度下为2000左右
        StrainTheoryRightBound = StrainTheoryRightBound + Sheets(StrainSheetName).Cells(3, i).Width
    Next
    
    '获取理论值图片指针
    DispTheoryShapeObjCounts = GetTheoryGraph(DispSheetName, DispTheoryShapeObjArray(), DispTheoryLeftBound, DispTheoryRightBound)
    StrainTheoryShapeObjCounts = GetTheoryGraph(StrainSheetName, StrainTheoryShapeObjArray(), StrainTheoryLeftBound, StrainTheoryRightBound)
     
    Dim tbl As Table
    Dim tableDataStartRow As Integer
    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\" & templateFileName, ThisWorkbook.Path & "\AutoCalcReportSource.docx"
    
    wordApp.Documents.Open ThisWorkbook.Path & "\AutoCalcReportSource.docx"

    wordApp.Visible = False
    
    tableDataStartRow = 3
    For i = 1 To NWCs                            'i定位工况
        
        '插入原始数据处理表标题
        wordApp.ActiveDocument.Variables(dispRawTbTitleVar(i)).value = dispRawTbTitle(i)
        
        '插入原始数据处理表格
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(dispRawTblBookmarks(i)).Range, NumRows:=DispUbound(i) + 1, NumColumns:=7) 'NumRows+1表示表头
        
        '不调表格宽度
        
        tbl.Cell(1, 1).Range.InsertAfter "测点号"
        tbl.Cell(1, 2).Range.InsertAfter "初始读数"
        tbl.Cell(1, 3).Range.InsertAfter "满载"
        tbl.Cell(1, 4).Range.InsertAfter "退载"
        tbl.Cell(1, 5).Range.InsertAfter "总挠度"
        tbl.Cell(1, 6).Range.InsertAfter "弹性挠度"
        tbl.Cell(1, 7).Range.InsertAfter "残余变形"

        tbl.Cell(1, 1).Range.Font.Bold = True
        tbl.Cell(1, 2).Range.Font.Bold = True
        tbl.Cell(1, 3).Range.Font.Bold = True
        tbl.Cell(1, 4).Range.Font.Bold = True
        tbl.Cell(1, 5).Range.Font.Bold = True
        tbl.Cell(1, 6).Range.Font.Bold = True
        tbl.Cell(1, 7).Range.Font.Bold = True
        
        For j = 1 To DispUbound(i)               'j定位测点
            tbl.Cell(1 + j, 1).Range.InsertAfter CStr(NodeName(i, j))
            tbl.Cell(1 + j, 2).Range.InsertAfter Format(InitDisp(i, j), "Fixed")
            tbl.Cell(1 + j, 3).Range.InsertAfter Format(FullLoadDisp(i, j), "Fixed")
            tbl.Cell(1 + j, 4).Range.InsertAfter Format(UnLoadDisp(i, j), "Fixed")
            tbl.Cell(1 + j, 5).Range.InsertAfter Format(TotalDisp(i, j), "Fixed")
            tbl.Cell(1 + j, 6).Range.InsertAfter Format(ElasticDisp(i, j), "Fixed")
            tbl.Cell(1 + j, 7).Range.InsertAfter Format(RemainDisp(i, j), "Fixed")
        Next
        
        SetTableBorder tbl
        SetTableAlignment tbl
    Next i
    
    tableDataStartRow = 3
    For i = 1 To NWCs                            'i定位工况
        '插入表格标题
        wordApp.ActiveDocument.Variables(dispTbTitleVar(i)).value = dispTbTitle(i)
        '插入表格
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(dispTblBookmarks(i)).Range, NumRows:=DispUbound(i) + tableDataStartRow - 1, NumColumns:=7) 'NumRows+1表示表头
               
        SetDispTableWidth tbl, DispUbound(i) + tableDataStartRow - 1
        
        tbl.Cell(1, 1).Merge tbl.Cell(2, 1)
        tbl.Cell(1, 2).Merge tbl.Cell(1, 4)
        tbl.Cell(1, 3).Merge tbl.Cell(2, 5)
        tbl.Cell(1, 4).Merge tbl.Cell(2, 6)
        tbl.Cell(1, 5).Merge tbl.Cell(2, 7)
        
        tbl.Cell(1, 1).Range.Font.Bold = True
        tbl.Cell(1, 2).Range.Font.Bold = True
        tbl.Cell(2, 2).Range.Font.Bold = True
        tbl.Cell(2, 3).Range.Font.Bold = True
        tbl.Cell(2, 4).Range.Font.Bold = True
        tbl.Cell(1, 3).Range.Font.Bold = True
        tbl.Cell(1, 4).Range.Font.Bold = True
        tbl.Cell(1, 5).Range.Font.Bold = True
        
        tbl.Cell(1, 1).Range.InsertAfter "测点号"
        tbl.Cell(1, 2).Range.InsertAfter "实测值(mm)"
        tbl.Cell(2, 2).Range.InsertAfter "总变形"
        tbl.Cell(2, 3).Range.InsertAfter "弹性变形"
        tbl.Cell(2, 4).Range.InsertAfter "残余变形"
        tbl.Cell(1, 3).Range.InsertAfter "满载理论值(mm)"
        tbl.Cell(1, 4).Range.InsertAfter "校验系数"
        tbl.Cell(1, 5).Range.InsertAfter "相对残余变形"
        
        For j = 1 To DispUbound(i)               'j定位测点
            tbl.Cell(tableDataStartRow - 1 + j, 1).Range.InsertAfter Format(NodeName(i, j), "Fixed")
            tbl.Cell(tableDataStartRow - 1 + j, 2).Range.InsertAfter Format(TotalDisp(i, j), "Fixed")
            tbl.Cell(tableDataStartRow - 1 + j, 3).Range.InsertAfter Format(ElasticDisp(i, j), "Fixed")
            tbl.Cell(tableDataStartRow - 1 + j, 4).Range.InsertAfter Format(RemainDisp(i, j), "Fixed")
            tbl.Cell(tableDataStartRow - 1 + j, 5).Range.InsertAfter Format(TheoryDisp(i, j), "Fixed")
            tbl.Cell(tableDataStartRow - 1 + j, 6).Range.InsertAfter Format(CheckoutCoff(i, j), "Fixed")
            tbl.Cell(tableDataStartRow - 1 + j, 7).Range.InsertAfter Format(RefRemainDisp(i, j), "Percent")
        Next
        
        SetTableBorder tbl
        SetTableAlignment tbl

        If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox_Calc") = xlOn Then '注：只能使用xlOn，不能使用True（尽管真值为1）
            wordApp.ActiveDocument.Variables(dispChartTitleVar(i)).value = dispChartTitle(i)
            DispChartObjArray(i).Copy
            wordApp.ActiveDocument.Bookmarks(dispChartBookmarks(i)).Range.Paste
        End If
        
        '位移理论值
        wordApp.ActiveDocument.Variables(dispTheoryShapeTitleVar(i)).value = dispTheoryShapeTitle(i)
        
        '如果截图少了，则不贴
        If DispTheoryShapeObjCounts = NWCs Then
            DispTheoryShapeObjArray(i).CopyPicture
            wordApp.ActiveDocument.Bookmarks(dispTheoryShapeBookmarks(i)).Range.Paste
        End If
    Next i


    tableDataStartRow = 3
    For i = 1 To strainNWCs                      'i定位工况
        '插入原始数据处理表格标题
        wordApp.ActiveDocument.Variables(strainRawTableTitleVar(i)).value = strainRawTbTitle(i)
        
        '插入原始数据处理表格
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(strainRawTableBookmarks(i)).Range, NumRows:=StrainUbound(i) + tableDataStartRow - 1, NumColumns:=14) 'NumRows+1表示表头
        
        SetStrainRawTableWidth tbl, StrainUbound(i) + tableDataStartRow - 1
        
        tbl.Cell(1, 1).Merge tbl.Cell(2, 1)
        tbl.Cell(1, 2).Merge tbl.Cell(1, 3)
        tbl.Cell(1, 3).Merge tbl.Cell(1, 4)
        tbl.Cell(1, 4).Merge tbl.Cell(1, 5)
        tbl.Cell(1, 5).Merge tbl.Cell(1, 6)
        tbl.Cell(1, 6).Merge tbl.Cell(1, 7)
        tbl.Cell(1, 7).Merge tbl.Cell(2, 12)
        tbl.Cell(1, 8).Merge tbl.Cell(2, 13)
        tbl.Cell(1, 9).Merge tbl.Cell(2, 14)
        
        tbl.Cell(1, 1).Range.InsertAfter "测点号"
        tbl.Cell(1, 2).Range.InsertAfter "初始读数"
        tbl.Cell(1, 3).Range.InsertAfter "满载"
        tbl.Cell(1, 4).Range.InsertAfter "退载"
        tbl.Cell(1, 5).Range.InsertAfter "加载"
        tbl.Cell(1, 6).Range.InsertAfter "退载"
        tbl.Cell(1, 7).Range.InsertAfter "总应变"
        tbl.Cell(1, 8).Range.InsertAfter "弹性应变"
        tbl.Cell(1, 9).Range.InsertAfter "残余应变"

        tbl.Cell(2, 2).Range.InsertAfter "模数R"
        tbl.Cell(2, 3).Range.InsertAfter "温度T"
        tbl.Cell(2, 4).Range.InsertAfter "模数R"
        tbl.Cell(2, 5).Range.InsertAfter "温度T"
        tbl.Cell(2, 6).Range.InsertAfter "模数R"
        tbl.Cell(2, 7).Range.InsertAfter "温度T"
        tbl.Cell(2, 8).Range.InsertAfter "ΔR"
        tbl.Cell(2, 9).Range.InsertAfter "ΔT"
        tbl.Cell(2, 10).Range.InsertAfter "ΔR"
        tbl.Cell(2, 11).Range.InsertAfter "ΔT"
        
        For j = 1 To 9
            tbl.Cell(1, j).Range.Font.Bold = True
        Next j

        For j = 2 To 11
            tbl.Cell(2, j).Range.Font.Bold = True
        Next j
        
            '参考应变表格----------开始----------
'        tbl.Cell(1, 1).Range.InsertAfter "测点号"
'        tbl.Cell(1, 2).Range.InsertAfter "实测值(με)"
'        tbl.Cell(2, 2).Range.InsertAfter "总应变"
'        tbl.Cell(2, 3).Range.InsertAfter "弹性应变"
'        tbl.Cell(2, 4).Range.InsertAfter "残余应变"
'        tbl.Cell(1, 3).Range.InsertAfter "满载理论值(με)"
'        tbl.Cell(1, 4).Range.InsertAfter "校验系数"
'        tbl.Cell(1, 5).Range.InsertAfter "相对残余应变"
'        For j = 1 To StrainUbound(i)             'j定位测点
'            tbl.Cell(tableDataStartRow - 1 + j, 1).Range.InsertAfter Format(StrainNodeName(i, j), "Fixed")
'            tbl.Cell(tableDataStartRow - 1 + j, 2).Range.InsertAfter INTTotalStrain(i, j)
'            tbl.Cell(tableDataStartRow - 1 + j, 3).Range.InsertAfter INTElasticStrain(i, j)
'            tbl.Cell(tableDataStartRow - 1 + j, 4).Range.InsertAfter INTRemainStrain(i, j)
'            tbl.Cell(tableDataStartRow - 1 + j, 5).Range.InsertAfter Round(TheoryStrain(i, j), 0)
'            tbl.Cell(tableDataStartRow - 1 + j, 6).Range.InsertAfter Format(INTDivStrainCheckoutCoff(i, j), "Fixed")
'            tbl.Cell(tableDataStartRow - 1 + j, 7).Range.InsertAfter Format(INTDivRefRemainStrain(i, j), "Percent")
'        Next
        '参考应变表格----------结束----------
        For j = 1 To StrainUbound(i)             'j定位测点
            tbl.Cell(tableDataStartRow - 1 + j, 1).Range.InsertAfter CStr((StrainNodeName(i, j)))
            tbl.Cell(tableDataStartRow - 1 + j, 2).Range.InsertAfter Format(InitStrainR0(i, j), "#0.0")
            tbl.Cell(tableDataStartRow - 1 + j, 3).Range.InsertAfter Format(InitStrainT0(i, j), "#0.0")
            tbl.Cell(tableDataStartRow - 1 + j, 4).Range.InsertAfter Format(FullLoadStrainR0(i, j), "#0.0")
            tbl.Cell(tableDataStartRow - 1 + j, 5).Range.InsertAfter Format(FullLoadStrainT0(i, j), "#0.0")
            tbl.Cell(tableDataStartRow - 1 + j, 6).Range.InsertAfter Format(UnLoadStrainR0(i, j), "#0.0")
            tbl.Cell(tableDataStartRow - 1 + j, 7).Range.InsertAfter Format(UnLoadStrainT0(i, j), "#0.0")
            tbl.Cell(tableDataStartRow - 1 + j, 8).Range.InsertAfter Format(FullLoadStrainR(i, j), "#0.0")
            tbl.Cell(tableDataStartRow - 1 + j, 9).Range.InsertAfter Format(FullLoadStrainT(i, j), "#0.0")
            tbl.Cell(tableDataStartRow - 1 + j, 10).Range.InsertAfter Format(UnLoadStrainR(i, j), "#0.0")
            tbl.Cell(tableDataStartRow - 1 + j, 11).Range.InsertAfter Format(UnLoadStrainT(i, j), "#0.0")
            tbl.Cell(tableDataStartRow - 1 + j, 12).Range.InsertAfter Round(TheoryStrain(i, j), 0)
            tbl.Cell(tableDataStartRow - 1 + j, 13).Range.InsertAfter INTElasticStrain(i, j)
            tbl.Cell(tableDataStartRow - 1 + j, 14).Range.InsertAfter INTRemainStrain(i, j)
        Next
'        For j = 1 To StrainUbound(i)             'j定位测点
'            tbl.Cell(tableDataStartRow - 1 + j, 1).Range.InsertAfter CStr((StrainNodeName(i, j)))
'            tbl.Cell(tableDataStartRow - 1 + j, 2).Range.InsertAfter Format(InitStrainR0(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 3).Range.InsertAfter Format(InitStrainT0(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 4).Range.InsertAfter Format(FullLoadStrainR0(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 5).Range.InsertAfter Format(FullLoadStrainT0(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 6).Range.InsertAfter Format(UnLoadStrainR0(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 7).Range.InsertAfter Format(UnLoadStrainT0(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 8).Range.InsertAfter Format(FullLoadStrainR(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 9).Range.InsertAfter Format(FullLoadStrainT(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 10).Range.InsertAfter Format(UnLoadStrainR(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 11).Range.InsertAfter Format(UnLoadStrainT(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 12).Range.InsertAfter Format(TotalStrain(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 13).Range.InsertAfter Format(ElasticStrain(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 14).Range.InsertAfter Format(RemainStrain(i, j), "#0.0")
'        Next
        
        SetTableBorder tbl
        SetTableAlignment tbl
    Next
    
    tableDataStartRow = 3
    For i = 1 To strainNWCs                      'i定位工况
        '插入表格标题
        wordApp.ActiveDocument.Variables(strainTbTitleVar(i)).value = strainTbTitle(i)
        '插入表格
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(strainTableBookmarks(i)).Range, NumRows:=StrainUbound(i) + tableDataStartRow - 1, NumColumns:=8) 'NumRows+1表示表头
        
        SetCalcStrainTableWidth tbl, StrainUbound(i) + tableDataStartRow - 1
        
        tbl.Cell(1, 1).Merge tbl.Cell(2, 1)
        tbl.Cell(1, 2).Merge tbl.Cell(1, 4)
        tbl.Cell(1, 3).Merge tbl.Cell(2, 5)
        tbl.Cell(1, 4).Merge tbl.Cell(2, 6)
        tbl.Cell(1, 5).Merge tbl.Cell(2, 7)
        tbl.Cell(1, 6).Merge tbl.Cell(2, 8)
        
        tbl.Cell(1, 1).Range.Font.Bold = True
        tbl.Cell(1, 2).Range.Font.Bold = True
        tbl.Cell(2, 2).Range.Font.Bold = True
        tbl.Cell(2, 3).Range.Font.Bold = True
        tbl.Cell(2, 4).Range.Font.Bold = True
        tbl.Cell(1, 3).Range.Font.Bold = True
        tbl.Cell(1, 4).Range.Font.Bold = True
        tbl.Cell(1, 5).Range.Font.Bold = True
        tbl.Cell(1, 6).Range.Font.Bold = True
        
        tbl.Cell(1, 1).Range.InsertAfter "测点号"
        tbl.Cell(1, 2).Range.InsertAfter "实测应变值(με)"
        tbl.Cell(2, 2).Range.InsertAfter "总应变"
        tbl.Cell(2, 3).Range.InsertAfter "弹性应变"
        tbl.Cell(2, 4).Range.InsertAfter "残余应变"
        tbl.Cell(1, 3).Range.InsertAfter "满载应力理论值" & vbCrLf & "（MPa）"
        tbl.Cell(1, 4).Range.InsertAfter "满载理论值" & vbCrLf & "(με)"
        tbl.Cell(1, 5).Range.InsertAfter "校验系数"
        tbl.Cell(1, 6).Range.InsertAfter "相对残余应变"

'        For j = 1 To StrainUbound(i)             'j定位测点
'            tbl.Cell(tableDataStartRow - 1 + j, 1).Range.InsertAfter CStr((StrainNodeName(i, j)))
'            tbl.Cell(tableDataStartRow - 1 + j, 2).Range.InsertAfter Format(TotalStrain(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 3).Range.InsertAfter Format(ElasticStrain(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 4).Range.InsertAfter Format(RemainStrain(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 5).Range.InsertAfter Format(TheoryStress(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 6).Range.InsertAfter Format(TheoryStrain(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 7).Range.InsertAfter Format(StrainCheckoutCoff(i, j), "Fixed")
'            tbl.Cell(tableDataStartRow - 1 + j, 8).Range.InsertAfter Format(RefRemainStrain(i, j), "Percent")
'        Next
        
        For j = 1 To StrainUbound(i)             'j定位测点
            tbl.Cell(tableDataStartRow - 1 + j, 1).Range.InsertAfter CStr((StrainNodeName(i, j)))
            tbl.Cell(tableDataStartRow - 1 + j, 2).Range.InsertAfter INTTotalStrain(i, j)
            tbl.Cell(tableDataStartRow - 1 + j, 3).Range.InsertAfter INTElasticStrain(i, j)
            tbl.Cell(tableDataStartRow - 1 + j, 4).Range.InsertAfter INTRemainStrain(i, j)
            tbl.Cell(tableDataStartRow - 1 + j, 5).Range.InsertAfter Format(TheoryStress(i, j), "#0.0")
            tbl.Cell(tableDataStartRow - 1 + j, 6).Range.InsertAfter Round(TheoryStrain(i, j), 0)
            tbl.Cell(tableDataStartRow - 1 + j, 7).Range.InsertAfter Format(INTDivStrainCheckoutCoff(i, j), "Fixed")
            tbl.Cell(tableDataStartRow - 1 + j, 8).Range.InsertAfter Format(INTDivRefRemainStrain(i, j), "Percent")
        Next

        '参考应变表格----------开始----------
'        tbl.Cell(1, 1).Range.InsertAfter "测点号"
'        tbl.Cell(1, 2).Range.InsertAfter "实测值(με)"
'        tbl.Cell(2, 2).Range.InsertAfter "总应变"
'        tbl.Cell(2, 3).Range.InsertAfter "弹性应变"
'        tbl.Cell(2, 4).Range.InsertAfter "残余应变"
'        tbl.Cell(1, 3).Range.InsertAfter "满载理论值(με)"
'        tbl.Cell(1, 4).Range.InsertAfter "校验系数"
'        tbl.Cell(1, 5).Range.InsertAfter "相对残余应变"
'        For j = 1 To StrainUbound(i)             'j定位测点
'            tbl.Cell(tableDataStartRow - 1 + j, 1).Range.InsertAfter Format(StrainNodeName(i, j), "Fixed")
'            tbl.Cell(tableDataStartRow - 1 + j, 2).Range.InsertAfter INTTotalStrain(i, j)
'            tbl.Cell(tableDataStartRow - 1 + j, 3).Range.InsertAfter INTElasticStrain(i, j)
'            tbl.Cell(tableDataStartRow - 1 + j, 4).Range.InsertAfter INTRemainStrain(i, j)
'            tbl.Cell(tableDataStartRow - 1 + j, 5).Range.InsertAfter Round(TheoryStrain(i, j), 0)
'            tbl.Cell(tableDataStartRow - 1 + j, 6).Range.InsertAfter Format(INTDivStrainCheckoutCoff(i, j), "Fixed")
'            tbl.Cell(tableDataStartRow - 1 + j, 7).Range.InsertAfter Format(INTDivRefRemainStrain(i, j), "Percent")
'        Next
        '参考应变表格----------结束----------
        
        SetTableBorder tbl
        SetTableAlignment tbl
        
        If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox_Calc") = xlOn Then '注：只能使用xlOn，不能使用True（尽管真值为1）
            wordApp.ActiveDocument.Variables(strainChartTitleVar(i)).value = strainChartTitle(i)
            StrainChartObjArray(i).Copy
            wordApp.ActiveDocument.Bookmarks(strainChartBookmarks(i)).Range.Paste
        End If
        
        '应变理论值
        wordApp.ActiveDocument.Variables(strainTheoryShapeTitleVar(i)).value = strainTheoryShapeTitle(i)
        
        '如果截图少了，则不贴
        If StrainTheoryShapeObjCounts = strainNWCs Then
            StrainTheoryShapeObjArray(i).CopyPicture
            wordApp.ActiveDocument.Bookmarks(strainTheoryShapeBookmarks(i)).Range.Paste
        End If
    Next
    
    wordApp.ActiveDocument.Fields.Update         '更新域
    
    DelCalcSpecifiedBookmarks wordApp.ActiveDocument
    UnlinkCalcSpecifiedDocVarFields wordApp.ActiveDocument

    '以防万一，新建标签（特别是在MS Word中）
    wordApp.ActiveDocument.Application.CaptionLabels.Add name:="图"
    wordApp.ActiveDocument.Application.CaptionLabels.Add name:="表"
    
    'TODO:根据报告模板修改参数
'    Dim tableOffset As Integer
'    Dim chartOffset As Integer
'    tableOffset = 1
'    chartOffset = 5
    
    '设置题注
    SetAutoCalcReportCaptions wordApp
    
    wordApp.Documents.Save
    
    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\" & calcFileName
    resultFlag = True
    
CloseWord:

    wordApp.Documents.Close
    wordApp.Quit
    Set wordApp = Nothing
    Set tbl = Nothing
    If resultFlag = True Then
        MsgBox ExportPromteString
    Else
        MsgBox "计算书导出失败！"
    End If
End Sub

'请先计算，再生成Word报告
Public Sub AutoReport()

    'Kill ThisWorkbook.Path & "\AutoReportSource.docx"
    GenAutoReportTemplate                        '先初始化模板
    
    Dim reportTemplateFileName As String
    reportTemplateFileName = "自动报告模板.docx"       '"报告模板.docx"

    Dim reportFileName As String
    reportFileName = "自动生成的报告.docx"
    
    
    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    
    Dim i, j As Integer
    Dim resultFlag As Boolean                    'True表示导出成功
    resultFlag = False
    
    Dim dispResultVar(1 To MAX_NWC) As String    '和word中响应DocVariable对应（常量数组）
    For i = 1 To MAX_NWC
        dispResultVar(i) = Replace("dispResult" & Str(i), " ", "")
    Next
    
    Dim dispSummaryVar(1 To MAX_NWC) As String   '和word中响应DocVariable对应（常量数组）
    For i = 1 To MAX_NWC
        dispSummaryVar(i) = Replace("dispSummary" & Str(i), " ", "")
    Next
    
    Dim dispTbTitleVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        dispTbTitleVar(i) = Replace("dispTbTitle" & Str(i), " ", "")
    Next
    Dim dispChartTitleVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        dispChartTitleVar(i) = Replace("dispChartTitle" & Str(i), " ", "")
    Next
    
    Dim strainResultVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        strainResultVar(i) = Replace("strainResult" & Str(i), " ", "")
    Next
    
    Dim strainSummaryVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        strainSummaryVar(i) = Replace("strainSummary" & CStr(i), " ", "")
    Next

    Dim strainTbTitleVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        strainTbTitleVar(i) = Replace("strainTbTitle" & CStr(i), " ", "")
    Next
    
    Dim strainChartTitleVar(1 To MAX_NWC) As String
    For i = 1 To MAX_NWC
        strainChartTitleVar(i) = Replace("strainChartTitle" & CStr(i), " ", "")
    Next
    
    Dim dispTblBookmarks(1 To MAX_NWC)
    For i = 1 To MAX_NWC
        dispTblBookmarks(i) = Replace("dispTable" & CStr(i), " ", "")
    Next
    
    Dim dispChartBookmarks(1 To MAX_NWC)
    For i = 1 To MAX_NWC
        dispChartBookmarks(i) = Replace("dispChart" & CStr(i), " ", "")
    Next
    
    Dim strainChartBookmarks(1 To MAX_NWC)
    For i = 1 To MAX_NWC
        strainChartBookmarks(i) = Replace("strainChart" & Str(i), " ", "")
    Next
    
    Dim strainTblBookmarks(1 To MAX_NWC)
    For i = 1 To MAX_NWC
        strainTblBookmarks(i) = Replace("strainTable" & Str(i), " ", "")
    Next

    Dim tbl As Table

    Dim tableDataStartRow As Integer             '表格数据起始行
    
    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\" & reportTemplateFileName, ThisWorkbook.Path & "\AutoReportSource.docx"
    
    wordApp.Documents.Open ThisWorkbook.Path & "\AutoReportSource.docx"

    wordApp.Visible = False
    
    tableDataStartRow = 3
    For i = 1 To NWCs                            'i定位工况
        wordApp.ActiveDocument.Variables(dispResultVar(i)).value = dispResult(i)
        If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOff Then '注：只能使用xlOn，不能使用True（尽管真值为1）
            dispSummary(i) = Replace(dispSummary(i), "，挠度实测值与理论计算值的关系曲线详见" & dispGraphCrossRef(i), "")
        End If
        wordApp.ActiveDocument.Variables(dispSummaryVar(i)).value = dispSummary(i)
        wordApp.ActiveDocument.Variables(dispTbTitleVar(i)).value = dispTbTitle(i)
        
        
        '插入表格
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(dispTblBookmarks(i)).Range, NumRows:=DispUbound(i) + tableDataStartRow - 1, NumColumns:=7) 'NumRows+x表示表头
        
        SetDispTableWidth tbl, DispUbound(i) + tableDataStartRow - 1    '设置各列宽度
               
        tbl.Cell(1, 1).Merge tbl.Cell(2, 1)
        tbl.Cell(1, 2).Merge tbl.Cell(1, 4)
        tbl.Cell(1, 3).Merge tbl.Cell(2, 5)
        tbl.Cell(1, 4).Merge tbl.Cell(2, 6)
        tbl.Cell(1, 5).Merge tbl.Cell(2, 7)
        
        tbl.Cell(1, 1).Range.InsertAfter "测点号"
        tbl.Cell(1, 2).Range.InsertAfter "实测值(mm)"
        tbl.Cell(2, 2).Range.InsertAfter "总变形"
        tbl.Cell(2, 3).Range.InsertAfter "弹性变形"
        tbl.Cell(2, 4).Range.InsertAfter "残余变形"
        tbl.Cell(1, 3).Range.InsertAfter "满载理论值(mm)"
        tbl.Cell(1, 4).Range.InsertAfter "校验系数"
        tbl.Cell(1, 5).Range.InsertAfter "相对残余变形"
        
        tbl.Cell(1, 1).Range.Font.Bold = True
        tbl.Cell(1, 2).Range.Font.Bold = True
        tbl.Cell(2, 2).Range.Font.Bold = True
        tbl.Cell(2, 3).Range.Font.Bold = True
        tbl.Cell(2, 4).Range.Font.Bold = True
        tbl.Cell(1, 3).Range.Font.Bold = True
        tbl.Cell(1, 4).Range.Font.Bold = True
        tbl.Cell(1, 5).Range.Font.Bold = True
        
        
        For j = 1 To DispUbound(i)               'j定位测点
            tbl.Cell(tableDataStartRow - 1 + j, 1).Range.InsertAfter Format(NodeName(i, j), "Fixed")
            
            tbl.Cell(tableDataStartRow - 1 + j, 2).Range.InsertAfter Format(TotalDisp(i, j), "Fixed")
            tbl.Cell(tableDataStartRow - 1 + j, 3).Range.InsertAfter Format(ElasticDisp(i, j), "Fixed")
            tbl.Cell(tableDataStartRow - 1 + j, 4).Range.InsertAfter Format(RemainDisp(i, j), "Fixed")
            tbl.Cell(tableDataStartRow - 1 + j, 5).Range.InsertAfter Format(TheoryDisp(i, j), "Fixed")
            tbl.Cell(tableDataStartRow - 1 + j, 6).Range.InsertAfter Format(CheckoutCoff(i, j), "Fixed")
            tbl.Cell(tableDataStartRow - 1 + j, 7).Range.InsertAfter Format(RefRemainDisp(i, j), "Percent")
        Next
        

        SetTableBorder tbl
        SetTableAlignment tbl
        
        '是否导出图表
        If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then '注：只能使用xlOn，不能使用True（尽管真值为1）
             wordApp.ActiveDocument.Variables(dispChartTitleVar(i)).value = dispChartTitle(i)
            DispChartObjArray(i).Copy
            wordApp.ActiveDocument.Bookmarks(dispChartBookmarks(i)).Range.Paste
        End If
        
    Next i
    tbl.Rows.Alignment = wdAlignRowCenter
    
    tableDataStartRow = 3
    For i = 1 To strainNWCs                      'i定位工况
      
        wordApp.ActiveDocument.Variables(strainResultVar(i)).value = strainResult(i)
        If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOff Then '注：只能使用xlOn，不能使用True（尽管真值为1）
            strainSummary(i) = Replace(strainSummary(i), "，应变实测值与理论计算值的关系曲线详见" & strainGraphCrossRef(i), "")
        End If
        wordApp.ActiveDocument.Variables(strainSummaryVar(i)).value = strainSummary(i)
        wordApp.ActiveDocument.Variables(strainTbTitleVar(i)).value = strainTbTitle(i)
        
        '插入表格
        
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(strainTblBookmarks(i)).Range, NumRows:=StrainUbound(i) + tableDataStartRow - 1, NumColumns:=7) 'NumRows+1表示表头
        
        SetStrainTableWidth tbl, StrainUbound(i) + tableDataStartRow - 1
        tbl.Cell(1, 1).Merge tbl.Cell(2, 1)
        tbl.Cell(1, 2).Merge tbl.Cell(1, 4)
        tbl.Cell(1, 3).Merge tbl.Cell(2, 5)
        tbl.Cell(1, 4).Merge tbl.Cell(2, 6)
        tbl.Cell(1, 5).Merge tbl.Cell(2, 7)
        
        tbl.Cell(1, 1).Range.InsertAfter "测点号"
        tbl.Cell(1, 2).Range.InsertAfter "实测值(με)"
        tbl.Cell(2, 2).Range.InsertAfter "总应变"
        tbl.Cell(2, 3).Range.InsertAfter "弹性应变"
        tbl.Cell(2, 4).Range.InsertAfter "残余应变"
        tbl.Cell(1, 3).Range.InsertAfter "满载理论值(με)"
        tbl.Cell(1, 4).Range.InsertAfter "校验系数"
        tbl.Cell(1, 5).Range.InsertAfter "相对残余应变"
        
        tbl.Cell(1, 1).Range.Font.Bold = True
        tbl.Cell(1, 2).Range.Font.Bold = True
        tbl.Cell(2, 2).Range.Font.Bold = True
        tbl.Cell(2, 3).Range.Font.Bold = True
        tbl.Cell(2, 4).Range.Font.Bold = True
        tbl.Cell(1, 3).Range.Font.Bold = True
        tbl.Cell(1, 4).Range.Font.Bold = True
        tbl.Cell(1, 5).Range.Font.Bold = True
        
        For j = 1 To StrainUbound(i)             'j定位测点
            tbl.Cell(tableDataStartRow - 1 + j, 1).Range.InsertAfter Format(StrainNodeName(i, j), "Fixed")
            tbl.Cell(tableDataStartRow - 1 + j, 2).Range.InsertAfter INTTotalStrain(i, j)
            tbl.Cell(tableDataStartRow - 1 + j, 3).Range.InsertAfter INTElasticStrain(i, j)
            tbl.Cell(tableDataStartRow - 1 + j, 4).Range.InsertAfter INTRemainStrain(i, j)
            tbl.Cell(tableDataStartRow - 1 + j, 5).Range.InsertAfter Round(TheoryStrain(i, j), 0)
            tbl.Cell(tableDataStartRow - 1 + j, 6).Range.InsertAfter Format(INTDivStrainCheckoutCoff(i, j), "Fixed")
            tbl.Cell(tableDataStartRow - 1 + j, 7).Range.InsertAfter Format(INTDivRefRemainStrain(i, j), "Percent")
        Next
        
        SetTableBorder tbl
        SetTableAlignment tbl
        
        '是否导出图表
        If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then '注：只能使用xlOn，不能使用True（尽管真值为1）
            wordApp.ActiveDocument.Variables(strainChartTitleVar(i)).value = strainChartTitle(i)
            StrainChartObjArray(i).Copy
            wordApp.ActiveDocument.Bookmarks(strainChartBookmarks(i)).Range.Paste
        End If
        
    Next i
   'wordApp.ActiveDocument.Tables(2).Rows.Alignment = wdAlignRowCenter
    
    wordApp.ActiveDocument.Fields.Update         '更新域

    'DelAllBookmarks wordApp.ActiveDocument       '删除所有书签
    DelSpecifiedBookmarks wordApp.ActiveDocument
    UnlinkSpecifiedDocVarFields wordApp.ActiveDocument
    'UnlinkAllFields wordApp.ActiveDocument       '解除所有域链接

    '以防万一，新建标签（特别是在MS Word中）
    wordApp.ActiveDocument.Application.CaptionLabels.Add name:="图"
    wordApp.ActiveDocument.Application.CaptionLabels.Add name:="表"
    
    'TODO:根据报告模板修改参数
    Dim tableOffset As Integer
    Dim chartOffset As Integer
    tableOffset = 1
    chartOffset = 5
    
    '设置题注
    SetAutoReportCaptions wordApp
    '设置交叉引用
    SetAutoReportCrossReferences wordApp, tableOffset, chartOffset
    
    SetReportCaptionsFormat wordApp
    SetReportMainBodyFormat wordApp
    wordApp.Documents.Save
    
    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\" & reportFileName
        
    resultFlag = True
CloseWord:
    wordApp.Documents.Close
    wordApp.Quit
    
    Set wordApp = Nothing
    Set tbl = Nothing
    
    If resultFlag = True Then
        MsgBox ExportPromteString
    Else
        MsgBox "报告导出失败！"
    End If
    
End Sub

'打开荷载试验word报告
Public Sub OpenReport()
    Dim wordApp As Word.Application
    Set wordApp = CreateObject("Word.Application")
    wordApp.Documents.Open fileName:=ThisWorkbook.Path & "\" & AutoReportFileName, ReadOnly:=False
    wordApp.Visible = True
    wordApp.Activate
End Sub

'打开荷载试验计算书报告
Public Sub OpenCalcReport()
    Dim wordApp As Word.Application
    Set wordApp = CreateObject("Word.Application")
    wordApp.Documents.Open fileName:=ThisWorkbook.Path & "\" & AutoCalcReportFileName, ReadOnly:=False
    wordApp.Visible = True
    wordApp.Activate
End Sub

'测试表格的生成操作
Public Sub testTable()
    Dim i, j, k As Integer

    Dim templateFileName As String               '模板文件
    templateFileName = "栏杆推力模板.docx"             '存放栏杆推力报告中所需要的资源
    
    Dim reportTemplateFileName As String
    reportTemplateFileName = "栏杆推力报告模板.docx"     '"报告模板.docx"
    
    Dim reportFileName As String
    reportFileName = "自动生成的栏杆推力报告test.docx"
    
    Dim tempFileName As String                   '临时文件的名称
    tempFileName = "temp.docx"
    
    Dim tempTemplateFileName As String           '临时模板文件的名称
    tempFileName = "temp.docx"
    
    Dim wordApp As Word.Application
    
    Dim tempDoc As Word.Document
    Dim templateDoc As New Word.Document
    
    Dim tbl As Word.Table                        'As Object    'Variant/Object/Range
    Dim tempTable As Word.Table
    
    Dim r As Word.Range
    
    
    Dim resultFlag As Boolean                    'True表示导出成功
    resultFlag = False
    
    Dim tbRowStart As Integer

    
    Dim tableDataStartRow As Integer             '表格数据起始行
    
    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\" & reportTemplateFileName, ThisWorkbook.Path & "\" & tempFileName
    FileCopy ThisWorkbook.Path & "\栏杆推力模板bak.docx", ThisWorkbook.Path & "\" & templateFileName
    
    Set tempDoc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & tempFileName)
    Set templateDoc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & templateFileName)

    wordApp.Visible = True

    
    tbRowStart = 2
    

    
    Set tbl = templateDoc.Tables(1)              '暂时仅支持分5级加载
    Set tempTable = tbl

    tempTable.Cell(tbRowStart + 1, 5).Range.InsertAfter Format(3#, "Percent")
    'Dim newRow
    'Set newRow = tempTable.Rows.Add(BeforeRow:=tempTable.Rows(1))
    'Set newRow = Nothing

    'InsertRowsBelow NumRows:=1
    tempTable.Cell(tbRowStart + 1, 5).Select
    'Selection.InsertRowsBelow NumRows:=1

    tempTable.Select
    wordApp.Selection.Copy
    tempDoc.Bookmarks("table1").Range.Paste

     
    tempDoc.Fields.Update                        '更新域
    
    tempDoc.Save
    
    tempDoc.SaveAs2 ThisWorkbook.Path & "\" & reportFileName
    resultFlag = True
CloseWord:

    wordApp.Documents.Close
    wordApp.Quit
    
    Set wordApp = Nothing
    Set tempDoc = Nothing
    Set templateDoc = Nothing
    'Set tbl = Nothing
    
    If resultFlag = True Then
        MsgBox "报告导出完成！"
    Else
        MsgBox "报告导出失败！"
    End If
    
End Sub

'TODO：仅支持10个测点
'请先计算，再生成报告（栏杆推力）
Public Sub AutoThrustReport()
    Dim i, j, k As Integer
    
    Dim thrustResult(1 To MAX_nThrust) As String '
    Dim thrustResultVar(1 To MAX_nThrust) As String
    For i = 1 To MAX_nThrust
        thrustResultVar(i) = Replace("thrustResult" & CStr(i), " ", "")
    Next
    
    Dim tableTitle(1 To MAX_nThrust) As String
    Dim tableTitleVar(1 To MAX_nThrust) As String
    For i = 1 To MAX_nThrust
        tableTitleVar(i) = Replace("tableTitle" & CStr(i), " ", "")
    Next
    
    
    Dim templateFileName As String               '模板文件
    templateFileName = "栏杆推力模板.docx"             '存放栏杆推力报告中所需要的资源
    
    Dim reportTemplateFileName As String
    reportTemplateFileName = "栏杆推力报告模板.docx"     '"报告模板.docx"
    
    Dim reportFileName As String
    reportFileName = "自动生成的栏杆推力报告.docx"
    
    Dim tempFileName As String                   '临时文件的名称
    tempFileName = "temp.docx"
    
    Dim tempTemplateFileName As String           '临时模板文件的名称
    tempFileName = "temp.docx"
    
    Dim tableBookmarks(1 To MAX_nThrust)
    For i = 1 To MAX_nThrust
        tableBookmarks(i) = Replace("table" & CStr(i), " ", "")
    Next
    
    Dim wordApp As Word.Application
    
    Dim tempDoc As Word.Document
    Dim templateDoc As New Word.Document
    
    Dim tbl As Word.Table                        'As Object    'Variant/Object/Range
    Dim tempTable As Word.Table
    
    Dim r As Word.Range
    
    
    Dim resultFlag As Boolean                    'True表示导出成功
    resultFlag = False
    
    Dim tbRowStart As Integer
    Dim tableDataStartRow As Integer             '表格数据起始行
    
    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\" & reportTemplateFileName, ThisWorkbook.Path & "\" & tempFileName
    FileCopy ThisWorkbook.Path & "\栏杆推力模板bak.docx", ThisWorkbook.Path & "\" & templateFileName
    
    Set tempDoc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & tempFileName)
    Set templateDoc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & templateFileName)

    wordApp.Visible = False
    
    tbRowStart = 2
    
    For i = 1 To nThrust
    
        '插入结果描述
        thrustResult(i) = "试验结果表明，在满载情况下，栏杆变形稳定，测点最大弹性变形值为" & Format(ThrustElasticDisp(i), "Fixed") & "mm，" _
                                                                                                  & "残余变形" & Format(ThrustRemainDisp(i), "Fixed") & "mm，" _
                                                                                                  & "相对残余变形为" & Format(ThrustRefRemainDisp(i), "Percent") & "。"
        
        tempDoc.Variables(thrustResultVar(i)).value = thrustResult(i)

        '插入表格标题
        tableTitle(i) = "表x-x " & CStr(i) & "#测点变形检测结果汇总表"
        tempDoc.Variables(tableTitleVar(i)).value = tableTitle(i)
    
    
        Set tbl = templateDoc.Tables(i)          '暂时仅支持分5级加载
        Set tempTable = tbl
        For k = 1 To ThrustLevel + 1
            tempTable.Cell(tbRowStart + k, 2).Range.InsertAfter Format(ThrustTotalDisp(i, k), "Fixed")

        Next k
        tempTable.Cell(tbRowStart + ThrustLevel, 3).Range.InsertAfter Format(ThrustElasticDisp(i), "Fixed")
        tempTable.Cell(tbRowStart + 1, 4).Range.InsertAfter Format(ThrustRemainDisp(i), "Fixed")
        tempTable.Cell(tbRowStart + 1, 5).Range.InsertAfter Format(ThrustRefRemainDisp(i), "Percent")
    
        'tempDoc.Bookmarks(tableBookmarks(i)).Range = tbl.Range
        'Set tempTable = tempDoc.Tables.Add(tableBookmarks(i).Range, NumRows:=1, NumColumns:=1)
        'Set tempTable = tbl
        tempTable.Select
        wordApp.Selection.Copy
        tempDoc.Bookmarks(tableBookmarks(i)).Range.Paste
       
    Next i
     
    tempDoc.Fields.Update                        '更新域

    DelAllBookmarks tempDoc                      '删除所有书签
    UnlinkAllFields tempDoc                      '解除所有域链接
    
    tempDoc.Save
    
    tempDoc.SaveAs2 ThisWorkbook.Path & "\" & reportFileName
    resultFlag = True
CloseWord:

    wordApp.Documents.Close
    wordApp.Quit
    
    Set wordApp = Nothing
    Set tempDoc = Nothing
    Set templateDoc = Nothing
    'Set tbl = Nothing
    
    If resultFlag = True Then
        MsgBox "报告导出完成！"
    Else
        MsgBox "报告导出失败！"
    End If
    
End Sub

'设置table的文字对齐方式
Private Sub SetTableAlignment(ByRef tbl As Table)
    On Error Resume Next
    tbl.Select
    tbl.Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    tbl.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    tbl.Range.Rows.Alignment = wdAlignRowCenter
End Sub

'设置table各个单元格宽度
Private Sub SetDispTableWidth(ByRef tbl As Table, ByVal NumRows As Integer)
    On Error Resume Next
    Dim NumColumns As Integer
    NumColumns = 7
    Dim WidthArray(1 To 7) As Single    '各列宽度
    WidthArray(1) = 52: WidthArray(2) = 51: WidthArray(3) = 60: WidthArray(4) = 60: WidthArray(5) = 77: WidthArray(6) = 62: WidthArray(7) = 62
    Dim i As Integer
    Dim j As Integer
    tbl.AllowAutoFit = False
    For i = 1 To NumColumns
        For j = 1 To NumRows
             tbl.Cell(j, i).SetWidth ColumnWidth:=WidthArray(i), RulerStyle:=wdAdjustFirstColumn
        Next j
    Next i

End Sub

'设置位移原始数据处理表各个单元格宽度
Private Sub SetStrainRawTableWidth(ByRef tbl As Table, ByVal NumRows As Integer)
    On Error Resume Next
    Dim NumColumns As Integer
    NumColumns = 14
    Dim WidthArray(1 To 14) As Single    '各列宽度
    WidthArray(1) = 35: WidthArray(2) = 38: WidthArray(3) = 38: WidthArray(4) = 38: WidthArray(5) = 38: WidthArray(6) = 38: WidthArray(7) = 38
    WidthArray(8) = 35: WidthArray(9) = 35: WidthArray(10) = 35: WidthArray(11) = 35: WidthArray(12) = 38: WidthArray(13) = 38: WidthArray(14) = 38
    Dim i As Integer
    Dim j As Integer
    tbl.AllowAutoFit = False
    For i = 1 To NumColumns
        For j = 1 To NumRows
             tbl.Cell(j, i).SetWidth ColumnWidth:=WidthArray(i), RulerStyle:=wdAdjustFirstColumn
        Next j
    Next i

End Sub

'设置table各个单元格宽度
Private Sub SetStrainTableWidth(ByRef tbl As Table, ByVal NumRows As Integer)
    On Error Resume Next
    Dim NumColumns As Integer
    NumColumns = 7
    Dim WidthArray(1 To 7) As Single    '各列宽度
    WidthArray(1) = 52: WidthArray(2) = 51: WidthArray(3) = 60: WidthArray(4) = 60: WidthArray(5) = 77: WidthArray(6) = 62: WidthArray(7) = 62
    Dim i As Integer
    Dim j As Integer
    tbl.AllowAutoFit = False
    For i = 1 To NumColumns
        For j = 1 To NumRows
             tbl.Cell(j, i).SetWidth ColumnWidth:=WidthArray(i), RulerStyle:=wdAdjustFirstColumn
        Next j
    Next i

End Sub

'设置计算书应变表格各个单元格宽度
Private Sub SetCalcStrainTableWidth(ByRef tbl As Table, ByVal NumRows As Integer)
    On Error Resume Next
    Dim NumColumns As Integer
    NumColumns = 8
    Dim WidthArray(1 To 8) As Single    '各列宽度
    WidthArray(1) = 52: WidthArray(2) = 51: WidthArray(3) = 60: WidthArray(4) = 60: WidthArray(5) = 77: WidthArray(6) = 77: WidthArray(7) = 62: WidthArray(8) = 62
    Dim i As Integer
    Dim j As Integer
    tbl.AllowAutoFit = False
    For i = 1 To NumColumns
        For j = 1 To NumRows
             tbl.Cell(j, i).SetWidth ColumnWidth:=WidthArray(i), RulerStyle:=wdAdjustFirstColumn
        Next j
    Next i

End Sub


'设置table的边界线
Private Sub SetTableBorder(ByRef tbl As Table)

    On Error Resume Next
    With tbl
        With .Borders
            .InsideLineStyle = wdLineStyleSingle
            .OutsideLineStyle = wdLineStyleSingle
        End With
    End With
    
    With tbl
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
        End With
    End With
End Sub

'解除指定的书签
Private Sub DelSpecifiedBookmarks(ByRef doc As Word.Document)

    DelArrayBookmarks doc, dispTblBookmarks
    DelArrayBookmarks doc, strainTblBookmarks
    
    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then '注：只能使用xlOn，不能使用True（尽管真值为1）
        DelArrayBookmarks doc, dispChartBookmarks
        DelArrayBookmarks doc, strainChartBookmarks
    End If
End Sub

'解除计算书指定的书签
Private Sub DelCalcSpecifiedBookmarks(ByRef doc As Word.Document)

    DelArrayBookmarks doc, dispRawTblBookmarks
    DelArrayBookmarks doc, dispTblBookmarks
    DelArrayBookmarks doc, dispTheoryShapeBookmarks
    DelArrayBookmarks doc, strainTblBookmarks
    DelArrayBookmarks doc, strainRawTableBookmarks
    DelArrayBookmarks doc, strainTheoryShapeBookmarks
    
    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox_Calc") = xlOn Then '注：只能使用xlOn，不能使用True（尽管真值为1）
        DelArrayBookmarks doc, dispChartBookmarks
        DelArrayBookmarks doc, strainChartBookmarks
    End If
End Sub

'将以一个字符串数组中任意元素命名的书签删除
Private Sub DelArrayBookmarks(ByRef doc As Word.Document, ByRef arr() As String)
    On Error Resume Next
    Dim i As Integer
    For i = 1 To UBound(arr)
        doc.Bookmarks(arr(i)).Delete
    Next i
End Sub

'删除所有书签
Private Sub DelAllBookmarks(ByRef doc As Word.Document)
    On Error Resume Next
    Dim bk As Bookmark
    For Each bk In doc.Bookmarks
        bk.Delete
    Next
End Sub

'将以一个字符串数组中任意元素命名的文档变量解除链接
Private Sub UnlinkArrayFields(ByRef doc As Word.Document, ByRef arr() As String)
    On Error Resume Next
    Dim i As Integer
    Dim f As Field
    For Each f In doc.Fields
        For i = 1 To UBound(arr)
            If f.result.Text = arr(i) Then
                f.Unlink
                Exit For
            End If
        Next i
    Next

End Sub

'解除指定的域（文档变量）链接
Private Sub UnlinkSpecifiedDocVarFields(ByRef doc As Word.Document)

    UnlinkArrayFields doc, dispSummary
    UnlinkArrayFields doc, dispResult
    UnlinkArrayFields doc, dispTbTitle

    UnlinkArrayFields doc, strainSummary
    UnlinkArrayFields doc, strainResult
    UnlinkArrayFields doc, strainTbTitle
    
    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then '注：只能使用xlOn，不能使用True（尽管真值为1）
        UnlinkArrayFields doc, dispChartTitle
        UnlinkArrayFields doc, strainChartTitle
    End If
End Sub

'解除指定的域（文档变量）链接
Private Sub UnlinkCalcSpecifiedDocVarFields(ByRef doc As Word.Document)

    UnlinkArrayFields doc, dispRawTbTitle
    UnlinkArrayFields doc, dispTbTitle
    UnlinkArrayFields doc, dispTheoryShapeTitle
    UnlinkArrayFields doc, strainRawTbTitle
    UnlinkArrayFields doc, strainTbTitle
    UnlinkArrayFields doc, strainTheoryShapeTitle
    
    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox_Calc") = xlOn Then '注：只能使用xlOn，不能使用True（尽管真值为1）
        UnlinkArrayFields doc, dispChartTitle
        UnlinkArrayFields doc, strainChartTitle
    End If
End Sub

'解除所有域链接
Private Sub UnlinkAllFields(ByRef doc As Word.Document)
    On Error Resume Next
    Dim f As Field
    For Each f In doc.Fields
        f.Unlink
    Next

End Sub


