Attribute VB_Name = "AutoWord"
Option Explicit
'���Ӵ���
'���Կ��У�����ǩ���������
'wordApp.ActiveDocument.Bookmarks("dispSummary1").Range.InsertAfter "111"
'���Բ��������
'wordApp.ActiveDocument.Tables.Add wordApp.ActiveDocument.Bookmarks("dispSummary1").Range, NumRows:=14 + 1, NumColumns:=7

'����ͳ�Ʋ�������
Public Const MaxElasticDeform_Index As Integer = 1
Public Const MinCheckoutCoff_Index As Integer = 2
Public Const MaxCheckoutCoff_Index As Integer = 3
Public Const MinRefRemainDeform_Index As Integer = 4
Public Const MaxRefRemainDeform_Index As Integer = 5

Public Const AutoReportFileName As String = "�Զ����ɵı���.docx" '�������鱨��
Public Const AutoCalcReportFileName As String = "�Զ����ɵļ�����.docx" '�������������

'�����Զ����ɵ�ģ�壨����ģ�塢������ģ�壩
Sub TestGenBookmarkAndDocVar()
    Dim resultFlag As Boolean
    resultFlag = False
    
    Dim templateFileName As String
    'templateFileName = "���Զ�����ģ��.docx"            '"����ģ��.docx"
    templateFileName = "�������������鱨��ģ��.docx"
    Dim fileName As String
    fileName = "���Զ�����.docx"
    
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
        r.Font.name = "����_GB2312"                '�ο����룺'Selection.Font.Name = "����_GB2312"
        r.ParagraphFormat.Alignment = wdAlignParagraphLeft
        '        .Selection.Text = "DOCVARIABLE dispResult1 \* MERGEFORMAT"
        '        .Selection.Range.Fields.Add .Selection.Range, wdFieldEmpty, , False
        '        .Selection.Font.Name = "����_GB2312"
        '
        '        r.MoveEnd , -1
        '
        .Selection.EndKey
        .Selection.TypeParagraph
        '.Selection.TypeText Text:="word"
        Set r = .Selection.Range
        r.Text = "DOCVARIABLE dispResult2 \* MERGEFORMAT"
        r.Fields.Add r, wdFieldEmpty, , False
        r.Font.name = "����_GB2312"                '�ο����룺'Selection.Font.Name = "����_GB2312"
        r.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.EndKey
        .Selection.TypeParagraph
        '        'Set r = Nothing
        '
        '        .Selection.Text = "DOCVARIABLE dispResult2 \* MERGEFORMAT"
        '        .Selection.Range.Fields.Add .Selection.Range, wdFieldEmpty, , False
        '        .Selection.Font.Name = "����_GB2312"
        
        '        .Selection.MoveEnd

        'Set r1 = .Selection.Range
        'r.MoveEnd , -1
        '         r.Text = "DOCVARIABLE dispResult2 \* MERGEFORMAT"
        '         r.Fields.Add r, wdFieldEmpty, , False
        '         r.Font.Name = "����_GB2312"    '�ο����룺'Selection.Font.Name = "����_GB2312"

        '�ο�����
        '.Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Name:="2"
        '.Selection.MoveDown Unit:=wdLine, Count:=5
        '.Selection.MoveRight Unit:=wdCharacter, Count:=9
        '.Selection.TypeText Text:="word"    '�ο�����
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
    'wordApp.ActiveDocument.Fields.Update    '������

    wordApp.Documents.Save
    
    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\" & fileName
    resultFlag = True
    
CloseWord:

    wordApp.Documents.Close
    wordApp.Quit
    Set wordApp = Nothing
    Set r = Nothing
    
    If resultFlag = True Then
        MsgBox "������ɣ�"
    Else
        MsgBox "����ʧ�ܣ�"
    End If
End Sub

'�����Զ����ɵ�ģ�壨����ģ�塢������ģ�壩
Sub testGenTemplate()
    Dim templateFileName As String
    templateFileName = "���Զ�����ģ��.docx"            '"����ģ��.docx"
    
    Dim fileName As String
    fileName = "���Զ�����.docx"
    
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
 
    wordApp.ActiveDocument.Fields.Update         '������

    wordApp.Documents.Save
    
    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\" & fileName
    resultFlag = True
    
CloseWord:

    wordApp.Documents.Close
    wordApp.Quit
    Set wordApp = Nothing
    Set tbl = Nothing
    If resultFlag = True Then
        MsgBox "������ɣ�"
    Else
        MsgBox "����ʧ�ܣ�"
    End If
End Sub

'���ñ���"����"��ʽ��ʽ
Public Sub SetReportMainBodyFormat(ByRef wordApp As Word.Application)
    With wordApp.Application.ActiveDocument.Styles.Item("����")
        With .Font
            .NameAscii = "Times New Roman"
            .NameOther = "Times New Roman"
            .NameFarEast = "����_GB2312"
            .Size = 12                           'С��
        End With
    End With
End Sub

'���ñ���"��ע"��ʽ��ʽ
Public Sub SetReportCaptionsFormat(ByRef wordApp As Word.Application)
    With wordApp.Application.ActiveDocument.Styles.Item("��ע")
        With .Font
            .NameAscii = "Times New Roman"
            .NameOther = "Times New Roman"
            .NameFarEast = "����_GB2312"
            .Size = 12                           'С��
        End With
    End With
End Sub

'�����Զ��������ע�ͽ�������
Public Sub SetAutoReportCaptions(ByRef wordApp As Word.Application)
    Dim i As Integer

    '�Ӷ�
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
            With .CaptionLabels("��")
                .NumberStyle = wdCaptionNumberStyleArabic
                .IncludeChapterNumber = -1
                .ChapterStyleLevel = 1
                .Separator = wdSeparatorHyphen
            End With
            .Selection.InsertCaption Label:="��", Title:="", TitleAutoText:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=False
            
        End With
    Next
    
    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then 'ע��ֻ��ʹ��xlOn������ʹ��True��������ֵΪ1��
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
                With .CaptionLabels("ͼ")
                    .NumberStyle = wdCaptionNumberStyleArabic
                    .IncludeChapterNumber = -1
                    .ChapterStyleLevel = 1
                    .Separator = wdSeparatorHyphen
                End With
                .Selection.InsertCaption Label:="ͼ", Title:="", TitleAutoText:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=False
            End With
        Next
    End If
    
    'Ӧ��
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
            With .CaptionLabels("��")
                .NumberStyle = wdCaptionNumberStyleArabic
                .IncludeChapterNumber = -1
                .ChapterStyleLevel = 1
                .Separator = wdSeparatorHyphen
            End With
            .Selection.InsertCaption Label:="��", Title:="", TitleAutoText:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=False
        End With
    Next
    
    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then 'ע��ֻ��ʹ��xlOn������ʹ��True��������ֵΪ1��
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
                With .CaptionLabels("ͼ")
                    .NumberStyle = wdCaptionNumberStyleArabic
                    .IncludeChapterNumber = -1
                    .ChapterStyleLevel = 1
                    .Separator = wdSeparatorHyphen
                End With
                .Selection.InsertCaption Label:="ͼ", Title:="", TitleAutoText:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=False
            End With
        Next
    End If
    

End Sub

'����ͼ������ע
'searchText:�����ַ�
'captionName:��ע��ǩ��
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

'���ü������Զ��������ע�ͽ�������
Public Sub SetAutoCalcReportCaptions(ByRef wordApp As Word.Application)
    Dim i As Integer
    '�Ӷ�
    For i = 1 To NWCs
        AddCaptions wordApp, dispRawTbTitle(i), "��"
        AddCaptions wordApp, dispTbTitle(i), "��"
    Next
    
    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox_Calc") = xlOn Then 'ע��ֻ��ʹ��xlOn������ʹ��True��������ֵΪ1��
        For i = 1 To NWCs
            AddCaptions wordApp, dispChartTitle(i), "ͼ"
        Next
    End If
    
    For i = 1 To NWCs
        AddCaptions wordApp, dispTheoryShapeTitle(i), "ͼ"
    Next
    
    'Ӧ��
    For i = 1 To strainNWCs
        AddCaptions wordApp, strainRawTbTitle(i), "��"
        AddCaptions wordApp, strainTbTitle(i), "��"
    Next
    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox_Calc") = xlOn Then 'ע��ֻ��ʹ��xlOn������ʹ��True��������ֵΪ1��
        For i = 1 To strainNWCs
               AddCaptions wordApp, strainChartTitle(i), "ͼ"
        Next
    End If
    
    For i = 1 To strainNWCs
        AddCaptions wordApp, strainTheoryShapeTitle(i), "ͼ"
    Next
    
End Sub

Public Sub SetAutoReportCrossReferences(ByRef wordApp As Word.Application, Optional tableOffset As Integer = 0, Optional chartOffset As Integer = 0)
    Dim i As Integer
    Dim rs As ReportService
    Set rs = New ReportService
    
    For i = 1 To NWCs
        With wordApp.Application
            '���뽻������
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
            .Selection.InsertCrossReference ReferenceType:="��", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=rs.GetDispCrossReferenceItem(i, GlobalWC, StrainGlobalWC, strainNWCs, tableOffset), _
        InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
        End With
        If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then 'ע��ֻ��ʹ��xlOn������ʹ��True��������ֵΪ1��
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
                .Selection.InsertCrossReference ReferenceType:="ͼ", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=rs.GetDispCrossReferenceItem(i, GlobalWC, StrainGlobalWC, strainNWCs, chartOffset), _
            InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
            End With
        End If
    Next
    
    For i = 1 To strainNWCs
        With wordApp.Application
            '���뽻������
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
            .Selection.InsertCrossReference ReferenceType:="��", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=rs.GetStrainCrossReferenceItem(i, GlobalWC, NWCs, StrainGlobalWC, tableOffset), _
        InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
        End With
        If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then 'ע��ֻ��ʹ��xlOn������ʹ��True��������ֵΪ1��
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
                .Selection.InsertCrossReference ReferenceType:="ͼ", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=rs.GetStrainCrossReferenceItem(i, GlobalWC, NWCs, StrainGlobalWC, chartOffset), _
            InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
            End With
        End If
    Next
    
    Set rs = Nothing
End Sub

'�����Զ�����ģ��
Public Sub GenAutoReportTemplate()
        
    Dim ax As ArrayService                       'ԭas��������ؼ���as��ͻ����Ϊas
    Set ax = New ArrayService
    
    Dim resultFlag As Boolean
    resultFlag = False
    
    Dim templateFileName As String
    templateFileName = "�������������鱨��ģ��.docx"           '"����ģ��.docx"
    
    Dim fileName As String
    fileName = "�Զ�����ģ��.docx"
    
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
    For i = 1 To NWCs                            'i��λ����
        With wordApp.ActiveDocument.Application
            Set r = .Selection.Range
            '        'r.MoveEnd , -1
    
            r.Text = "dispResult" & CStr(i)
            'r.Font.Name = "Times New Roman"            '�ο����룺'Selection.Font.Name = "����_GB2312"
            r.ParagraphFormat.Alignment = wdAlignParagraphLeft
            r.ParagraphFormat.CharacterUnitFirstLineIndent = 2
            r.Fields.Add r, wdFieldDocVariable, r.Text, True
            '        .Selection.Text = "DOCVARIABLE dispResult1 \* MERGEFORMAT"
            '        .Selection.Range.Fields.Add .Selection.Range, wdFieldEmpty, , False
            '        .Selection.Font.Name = "����_GB2312"
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
            'r.Font.Name = "����_GB2312"
            r.ParagraphFormat.Alignment = wdAlignParagraphLeft
            r.ParagraphFormat.CharacterUnitFirstLineIndent = 2
            r.Fields.Add r, wdFieldDocVariable, r.Text, True
            
            .Selection.EndKey
            .Selection.TypeParagraph

        End With
    Next i
    
    wordApp.ActiveDocument.Application.Selection.TypeParagraph
    
    wordApp.ActiveDocument.Application.Selection.GoTo what:=wdGoToBookmark, name:=ReportStartBookmarkName
    '���㱨�湤������
    Dim totalNWC As Integer                      '�ܹ�����
    totalNWC = ax.ArrayMax(GlobalWC)
    If totalNWC < ax.ArrayMax(StrainGlobalWC) Then '�Ƚ��Ӷȡ�Ӧ����󹤿�
        totalNWC = ax.ArrayMax(StrainGlobalWC)
    End If
    'Debug.Print totalNWC
    
    For i = 1 To totalNWC
        With wordApp.ActiveDocument.Application
            '�����Ӷ����й����������湤����Ӧ����д��
            For j = 1 To NWCs
                If GlobalWC(j) = i Then
                    Set r = .Selection.Range
                        
                    r.Text = "dispSummary" & CStr(j)
                    'r.Font.Name = "����_GB2312"
                    r.ParagraphFormat.Alignment = wdAlignParagraphLeft
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 2
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                        
                    .Selection.EndKey
                    .Selection.TypeParagraph
                        
                    Set r = .Selection.Range
                    r.Text = "dispTbTitle" & CStr(j)
                    'r.Font.Name = "����_GB2312"
                    r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                    
                    .Selection.EndKey
                    .Selection.TypeParagraph
                        
                    '�Ӷȱ���ǩ
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
                        '�Ӷ�ͼ��ǩ
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
                        
                        '�Ӷ�ͼ����
                        Set r = .Selection.Range
                        r.Text = "dispChartTitle" & CStr(j)
                        'r.Font.Name = "����_GB2312"
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
                    'r.Font.Name = "����_GB2312"
                    r.ParagraphFormat.Alignment = wdAlignParagraphLeft
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 2
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                        
                    .Selection.EndKey
                    .Selection.TypeParagraph
                        
                    Set r = .Selection.Range
                    r.Text = "strainTbTitle" & CStr(j)
                    'r.Font.Name = "����_GB2312"
                    r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                         
                    .Selection.EndKey
                    .Selection.TypeParagraph
                    'Ӧ�����ǩ
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
                        'Ӧ��ͼ��ǩ
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
                        'r.Font.Name = "����_GB2312"
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
    'wordApp.ActiveDocument.Fields.Update    '������

    wordApp.Documents.Save
    
    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\" & fileName
    resultFlag = True
    
CloseWord:

    wordApp.Documents.Close
    wordApp.Quit
    Set wordApp = Nothing
    Set r = Nothing
    Set ax = Nothing
    
    'TODO��ģ������ʧ�ܣ�����ֹ����
    '    If resultFlag = True Then
    '        MsgBox "������ɣ�"
    '    Else
    '        MsgBox "����ʧ�ܣ�"
    '    End If
End Sub

'GenAutoCalcReportTemplate���������þ�������
Private Sub SetAutoCalcReportTemplateDetail(ByRef wordApp As Word.Application, ByVal totalNWC As Integer)
    
    Const CalcReportStartBookmarkName As String = "CalcReportStart"
    Dim r As Word.Range
    Dim i As Integer
    Dim j As Integer
    wordApp.Visible = False
    wordApp.ActiveDocument.Application.Selection.GoTo what:=wdGoToBookmark, name:=CalcReportStartBookmarkName
    For i = 1 To totalNWC
        With wordApp.ActiveDocument.Application
            '�����Ӷ����й����������湤����Ӧ����д��
            For j = 1 To NWCs
                If GlobalWC(j) = i Then
                    '�Ӷ�ԭʼ���ݴ�������
                    Set r = .Selection.Range
                    r.Text = "dispRawTbTitle" & CStr(j)
                    'r.Font.Name = "����_GB2312"
                    r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                    .Selection.EndKey
                    .Selection.TypeParagraph
                        
                    '�Ӷ�ԭʼ���ݴ������ǩ
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
                    
                    '�Ӷȱ����
                    Set r = .Selection.Range
                    r.Text = "dispTbTitle" & CStr(j)
                    'r.Font.Name = "����_GB2312"
                    r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                    .Selection.EndKey
                    .Selection.TypeParagraph
                        
                    '�Ӷȱ���ǩ
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
                        '�Ӷ�ͼ��ǩ
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
                    
                        '�Ӷ�ͼ����
                        Set r = .Selection.Range
                        r.Text = "dispChartTitle" & CStr(j)
                        'r.Font.Name = "����_GB2312"
                        r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                        r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                        r.Fields.Add r, wdFieldDocVariable, r.Text, True
                        .Selection.EndKey
                        .Selection.TypeParagraph

                    End If
                     '�Ӷ�����ͼ��ǩ
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
                    '�Ӷ�����ֵ����
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
                    'Ӧ��ԭʼ���ݴ�������
                    Set r = .Selection.Range
                    r.Text = "strainRawTbTitle" & CStr(j)
                    'r.Font.Name = "����_GB2312"
                    r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                    .Selection.EndKey
                    .Selection.TypeParagraph
                    'Ӧ��ԭʼ���ݴ������ǩ
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
                    'Ӧ������
                    Set r = .Selection.Range
                    r.Text = "strainTbTitle" & CStr(j)
                    'r.Font.Name = "����_GB2312"
                    r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                    r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                    r.Fields.Add r, wdFieldDocVariable, r.Text, True
                    .Selection.EndKey
                    .Selection.TypeParagraph
                    'Ӧ�����ǩ
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
                        'Ӧ��ͼ��ǩ
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
                        'Ӧ��ͼ����
                        Set r = .Selection.Range
                        r.Text = "strainChartTitle" & CStr(j)
                        'r.Font.Name = "����_GB2312"
                        r.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        r.ParagraphFormat.CharacterUnitFirstLineIndent = 0
                        r.ParagraphFormat.FirstLineIndent = wordApp.ActiveDocument.Application.CentimetersToPoints(0)
                        r.Fields.Add r, wdFieldDocVariable, r.Text, True
                         
                        .Selection.EndKey
                        .Selection.TypeParagraph
                    End If
                     'Ӧ������ͼ��ǩ
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
                    'Ӧ������ֵ����
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

'�����Զ�����ģ�壨��
Public Sub GenAutoCalcReportTemplate()
        
    Dim ax As ArrayService                       'ԭas��������ؼ���as��ͻ����Ϊas
    Set ax = New ArrayService
    
    Dim resultFlag As Boolean
    resultFlag = False
    
    Dim templateFileName As String
    'templateFileName = "���Զ�������ģ��.docx"           '"����ģ��.docx"
    templateFileName = "�������������������ģ��.docx"
    Dim fileName As String
    fileName = "�Զ�������ģ��.docx"
    
    Dim wordApp As Word.Application
    Dim doc As Word.Document

    Dim r1 As Word.Range
        
    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    FileCopy ThisWorkbook.Path & "\" & templateFileName, ThisWorkbook.Path & "\temp.docx"
    
    wordApp.Documents.Open ThisWorkbook.Path & "\temp.docx"
    'wordApp.ActiveDocument.Application.Selection.TypeParagraph
    
    '���㱨�湤������
    Dim totalNWC As Integer                      '�ܹ�����
    totalNWC = ax.ArrayMax(GlobalWC)
    If totalNWC < ax.ArrayMax(StrainGlobalWC) Then '�Ƚ��Ӷȡ�Ӧ����󹤿�
        totalNWC = ax.ArrayMax(StrainGlobalWC)
    End If
    
    SetAutoCalcReportTemplateDetail wordApp, totalNWC
    ' wordApp.ActiveDocument.Application.Selection
    ' wordApp.ActiveDocument.Selection.MoveDown 4, 1
    'wordApp.ActiveDocument.Fields.Update    '������

    wordApp.Documents.Save
    
    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\" & fileName
    resultFlag = True
    
CloseWord:

    wordApp.Documents.Close
    wordApp.Quit
    Set wordApp = Nothing
    Set ax = Nothing
    
    'TODO��ģ������ʧ�ܣ�����ֹ����
    '    If resultFlag = True Then
    '        MsgBox "������ɣ�"
    '    Else
    '        MsgBox "����ʧ�ܣ�"
    '    End If
End Sub
'��ȡ����ͼ��ָ��
'�㷨�����ͼƬ��߼��ڱ߽�֮�䣬��洢��ָ��
'sheetName:��ǩ��
'shapeObjArray():ͼƬ���飨������Ԥ���㹻��С��
'leftBoundry:��߽�
'rightBoundry:�ұ߽�
'����ֵ��ͼƬ����
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
    '��С��������
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


'�Զ����ɼ�����
Public Sub AutoCalcReport()
     
    GenAutoCalcReportTemplate
    Dim templateFileName As String
    templateFileName = "�Զ�������ģ��.docx"            '"����ģ��.docx"
    
    Dim calcFileName As String
    calcFileName = "�Զ����ɵļ�����.docx"

    Dim wordApp As Word.Application
    Dim doc As Word.Document
        
    Dim i, j As Integer
    
    Dim resultFlag As Boolean                    'True��ʾ�����ɹ�
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
    
    '�����ͼ���߽�
    Dim DispTheoryLeftBound As Integer     '��ͼ��߽�
    Dim DispTheoryRightBound As Integer      '��ͼ�ұ߽�
    DispTheoryLeftBound = 0    'Ĭ�ϱ������Ϊ1200����
    For i = 1 To 23
        DispTheoryLeftBound = DispTheoryLeftBound + Sheets(DispSheetName).Cells(3, i).Width
    Next
    DispTheoryRightBound = 0
    For i = 1 To 28    'Ĭ�ϱ������Ϊ1500����
        DispTheoryRightBound = DispTheoryRightBound + Sheets(DispSheetName).Cells(3, i).Width
    Next
    
    Dim StrainTheoryLeftBound As Integer     '��ͼ��߽�
    Dim StrainTheoryRightBound As Integer      '��ͼ�ұ߽�
    StrainTheoryLeftBound = 0    'Ĭ�ϱ������Ϊ1800����
    For i = 1 To 33
        StrainTheoryLeftBound = StrainTheoryLeftBound + Sheets(StrainSheetName).Cells(3, i).Width
    Next
    StrainTheoryRightBound = 0
    For i = 1 To 38    'Ĭ�ϱ������Ϊ2000����
        StrainTheoryRightBound = StrainTheoryRightBound + Sheets(StrainSheetName).Cells(3, i).Width
    Next
    
    '��ȡ����ֵͼƬָ��
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
    For i = 1 To NWCs                            'i��λ����
        
        '����ԭʼ���ݴ�������
        wordApp.ActiveDocument.Variables(dispRawTbTitleVar(i)).value = dispRawTbTitle(i)
        
        '����ԭʼ���ݴ�����
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(dispRawTblBookmarks(i)).Range, NumRows:=DispUbound(i) + 1, NumColumns:=7) 'NumRows+1��ʾ��ͷ
        
        '���������
        
        tbl.Cell(1, 1).Range.InsertAfter "����"
        tbl.Cell(1, 2).Range.InsertAfter "��ʼ����"
        tbl.Cell(1, 3).Range.InsertAfter "����"
        tbl.Cell(1, 4).Range.InsertAfter "����"
        tbl.Cell(1, 5).Range.InsertAfter "���Ӷ�"
        tbl.Cell(1, 6).Range.InsertAfter "�����Ӷ�"
        tbl.Cell(1, 7).Range.InsertAfter "�������"

        tbl.Cell(1, 1).Range.Font.Bold = True
        tbl.Cell(1, 2).Range.Font.Bold = True
        tbl.Cell(1, 3).Range.Font.Bold = True
        tbl.Cell(1, 4).Range.Font.Bold = True
        tbl.Cell(1, 5).Range.Font.Bold = True
        tbl.Cell(1, 6).Range.Font.Bold = True
        tbl.Cell(1, 7).Range.Font.Bold = True
        
        For j = 1 To DispUbound(i)               'j��λ���
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
    For i = 1 To NWCs                            'i��λ����
        '���������
        wordApp.ActiveDocument.Variables(dispTbTitleVar(i)).value = dispTbTitle(i)
        '������
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(dispTblBookmarks(i)).Range, NumRows:=DispUbound(i) + tableDataStartRow - 1, NumColumns:=7) 'NumRows+1��ʾ��ͷ
               
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
        
        tbl.Cell(1, 1).Range.InsertAfter "����"
        tbl.Cell(1, 2).Range.InsertAfter "ʵ��ֵ(mm)"
        tbl.Cell(2, 2).Range.InsertAfter "�ܱ���"
        tbl.Cell(2, 3).Range.InsertAfter "���Ա���"
        tbl.Cell(2, 4).Range.InsertAfter "�������"
        tbl.Cell(1, 3).Range.InsertAfter "��������ֵ(mm)"
        tbl.Cell(1, 4).Range.InsertAfter "У��ϵ��"
        tbl.Cell(1, 5).Range.InsertAfter "��Բ������"
        
        For j = 1 To DispUbound(i)               'j��λ���
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

        If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox_Calc") = xlOn Then 'ע��ֻ��ʹ��xlOn������ʹ��True��������ֵΪ1��
            wordApp.ActiveDocument.Variables(dispChartTitleVar(i)).value = dispChartTitle(i)
            DispChartObjArray(i).Copy
            wordApp.ActiveDocument.Bookmarks(dispChartBookmarks(i)).Range.Paste
        End If
        
        'λ������ֵ
        wordApp.ActiveDocument.Variables(dispTheoryShapeTitleVar(i)).value = dispTheoryShapeTitle(i)
        
        '�����ͼ���ˣ�����
        If DispTheoryShapeObjCounts = NWCs Then
            DispTheoryShapeObjArray(i).CopyPicture
            wordApp.ActiveDocument.Bookmarks(dispTheoryShapeBookmarks(i)).Range.Paste
        End If
    Next i


    tableDataStartRow = 3
    For i = 1 To strainNWCs                      'i��λ����
        '����ԭʼ���ݴ��������
        wordApp.ActiveDocument.Variables(strainRawTableTitleVar(i)).value = strainRawTbTitle(i)
        
        '����ԭʼ���ݴ�����
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(strainRawTableBookmarks(i)).Range, NumRows:=StrainUbound(i) + tableDataStartRow - 1, NumColumns:=14) 'NumRows+1��ʾ��ͷ
        
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
        
        tbl.Cell(1, 1).Range.InsertAfter "����"
        tbl.Cell(1, 2).Range.InsertAfter "��ʼ����"
        tbl.Cell(1, 3).Range.InsertAfter "����"
        tbl.Cell(1, 4).Range.InsertAfter "����"
        tbl.Cell(1, 5).Range.InsertAfter "����"
        tbl.Cell(1, 6).Range.InsertAfter "����"
        tbl.Cell(1, 7).Range.InsertAfter "��Ӧ��"
        tbl.Cell(1, 8).Range.InsertAfter "����Ӧ��"
        tbl.Cell(1, 9).Range.InsertAfter "����Ӧ��"

        tbl.Cell(2, 2).Range.InsertAfter "ģ��R"
        tbl.Cell(2, 3).Range.InsertAfter "�¶�T"
        tbl.Cell(2, 4).Range.InsertAfter "ģ��R"
        tbl.Cell(2, 5).Range.InsertAfter "�¶�T"
        tbl.Cell(2, 6).Range.InsertAfter "ģ��R"
        tbl.Cell(2, 7).Range.InsertAfter "�¶�T"
        tbl.Cell(2, 8).Range.InsertAfter "��R"
        tbl.Cell(2, 9).Range.InsertAfter "��T"
        tbl.Cell(2, 10).Range.InsertAfter "��R"
        tbl.Cell(2, 11).Range.InsertAfter "��T"
        
        For j = 1 To 9
            tbl.Cell(1, j).Range.Font.Bold = True
        Next j

        For j = 2 To 11
            tbl.Cell(2, j).Range.Font.Bold = True
        Next j
        
            '�ο�Ӧ����----------��ʼ----------
'        tbl.Cell(1, 1).Range.InsertAfter "����"
'        tbl.Cell(1, 2).Range.InsertAfter "ʵ��ֵ(�̦�)"
'        tbl.Cell(2, 2).Range.InsertAfter "��Ӧ��"
'        tbl.Cell(2, 3).Range.InsertAfter "����Ӧ��"
'        tbl.Cell(2, 4).Range.InsertAfter "����Ӧ��"
'        tbl.Cell(1, 3).Range.InsertAfter "��������ֵ(�̦�)"
'        tbl.Cell(1, 4).Range.InsertAfter "У��ϵ��"
'        tbl.Cell(1, 5).Range.InsertAfter "��Բ���Ӧ��"
'        For j = 1 To StrainUbound(i)             'j��λ���
'            tbl.Cell(tableDataStartRow - 1 + j, 1).Range.InsertAfter Format(StrainNodeName(i, j), "Fixed")
'            tbl.Cell(tableDataStartRow - 1 + j, 2).Range.InsertAfter INTTotalStrain(i, j)
'            tbl.Cell(tableDataStartRow - 1 + j, 3).Range.InsertAfter INTElasticStrain(i, j)
'            tbl.Cell(tableDataStartRow - 1 + j, 4).Range.InsertAfter INTRemainStrain(i, j)
'            tbl.Cell(tableDataStartRow - 1 + j, 5).Range.InsertAfter Round(TheoryStrain(i, j), 0)
'            tbl.Cell(tableDataStartRow - 1 + j, 6).Range.InsertAfter Format(INTDivStrainCheckoutCoff(i, j), "Fixed")
'            tbl.Cell(tableDataStartRow - 1 + j, 7).Range.InsertAfter Format(INTDivRefRemainStrain(i, j), "Percent")
'        Next
        '�ο�Ӧ����----------����----------
        For j = 1 To StrainUbound(i)             'j��λ���
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
'        For j = 1 To StrainUbound(i)             'j��λ���
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
    For i = 1 To strainNWCs                      'i��λ����
        '���������
        wordApp.ActiveDocument.Variables(strainTbTitleVar(i)).value = strainTbTitle(i)
        '������
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(strainTableBookmarks(i)).Range, NumRows:=StrainUbound(i) + tableDataStartRow - 1, NumColumns:=8) 'NumRows+1��ʾ��ͷ
        
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
        
        tbl.Cell(1, 1).Range.InsertAfter "����"
        tbl.Cell(1, 2).Range.InsertAfter "ʵ��Ӧ��ֵ(�̦�)"
        tbl.Cell(2, 2).Range.InsertAfter "��Ӧ��"
        tbl.Cell(2, 3).Range.InsertAfter "����Ӧ��"
        tbl.Cell(2, 4).Range.InsertAfter "����Ӧ��"
        tbl.Cell(1, 3).Range.InsertAfter "����Ӧ������ֵ" & vbCrLf & "��MPa��"
        tbl.Cell(1, 4).Range.InsertAfter "��������ֵ" & vbCrLf & "(�̦�)"
        tbl.Cell(1, 5).Range.InsertAfter "У��ϵ��"
        tbl.Cell(1, 6).Range.InsertAfter "��Բ���Ӧ��"

'        For j = 1 To StrainUbound(i)             'j��λ���
'            tbl.Cell(tableDataStartRow - 1 + j, 1).Range.InsertAfter CStr((StrainNodeName(i, j)))
'            tbl.Cell(tableDataStartRow - 1 + j, 2).Range.InsertAfter Format(TotalStrain(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 3).Range.InsertAfter Format(ElasticStrain(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 4).Range.InsertAfter Format(RemainStrain(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 5).Range.InsertAfter Format(TheoryStress(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 6).Range.InsertAfter Format(TheoryStrain(i, j), "#0.0")
'            tbl.Cell(tableDataStartRow - 1 + j, 7).Range.InsertAfter Format(StrainCheckoutCoff(i, j), "Fixed")
'            tbl.Cell(tableDataStartRow - 1 + j, 8).Range.InsertAfter Format(RefRemainStrain(i, j), "Percent")
'        Next
        
        For j = 1 To StrainUbound(i)             'j��λ���
            tbl.Cell(tableDataStartRow - 1 + j, 1).Range.InsertAfter CStr((StrainNodeName(i, j)))
            tbl.Cell(tableDataStartRow - 1 + j, 2).Range.InsertAfter INTTotalStrain(i, j)
            tbl.Cell(tableDataStartRow - 1 + j, 3).Range.InsertAfter INTElasticStrain(i, j)
            tbl.Cell(tableDataStartRow - 1 + j, 4).Range.InsertAfter INTRemainStrain(i, j)
            tbl.Cell(tableDataStartRow - 1 + j, 5).Range.InsertAfter Format(TheoryStress(i, j), "#0.0")
            tbl.Cell(tableDataStartRow - 1 + j, 6).Range.InsertAfter Round(TheoryStrain(i, j), 0)
            tbl.Cell(tableDataStartRow - 1 + j, 7).Range.InsertAfter Format(INTDivStrainCheckoutCoff(i, j), "Fixed")
            tbl.Cell(tableDataStartRow - 1 + j, 8).Range.InsertAfter Format(INTDivRefRemainStrain(i, j), "Percent")
        Next

        '�ο�Ӧ����----------��ʼ----------
'        tbl.Cell(1, 1).Range.InsertAfter "����"
'        tbl.Cell(1, 2).Range.InsertAfter "ʵ��ֵ(�̦�)"
'        tbl.Cell(2, 2).Range.InsertAfter "��Ӧ��"
'        tbl.Cell(2, 3).Range.InsertAfter "����Ӧ��"
'        tbl.Cell(2, 4).Range.InsertAfter "����Ӧ��"
'        tbl.Cell(1, 3).Range.InsertAfter "��������ֵ(�̦�)"
'        tbl.Cell(1, 4).Range.InsertAfter "У��ϵ��"
'        tbl.Cell(1, 5).Range.InsertAfter "��Բ���Ӧ��"
'        For j = 1 To StrainUbound(i)             'j��λ���
'            tbl.Cell(tableDataStartRow - 1 + j, 1).Range.InsertAfter Format(StrainNodeName(i, j), "Fixed")
'            tbl.Cell(tableDataStartRow - 1 + j, 2).Range.InsertAfter INTTotalStrain(i, j)
'            tbl.Cell(tableDataStartRow - 1 + j, 3).Range.InsertAfter INTElasticStrain(i, j)
'            tbl.Cell(tableDataStartRow - 1 + j, 4).Range.InsertAfter INTRemainStrain(i, j)
'            tbl.Cell(tableDataStartRow - 1 + j, 5).Range.InsertAfter Round(TheoryStrain(i, j), 0)
'            tbl.Cell(tableDataStartRow - 1 + j, 6).Range.InsertAfter Format(INTDivStrainCheckoutCoff(i, j), "Fixed")
'            tbl.Cell(tableDataStartRow - 1 + j, 7).Range.InsertAfter Format(INTDivRefRemainStrain(i, j), "Percent")
'        Next
        '�ο�Ӧ����----------����----------
        
        SetTableBorder tbl
        SetTableAlignment tbl
        
        If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox_Calc") = xlOn Then 'ע��ֻ��ʹ��xlOn������ʹ��True��������ֵΪ1��
            wordApp.ActiveDocument.Variables(strainChartTitleVar(i)).value = strainChartTitle(i)
            StrainChartObjArray(i).Copy
            wordApp.ActiveDocument.Bookmarks(strainChartBookmarks(i)).Range.Paste
        End If
        
        'Ӧ������ֵ
        wordApp.ActiveDocument.Variables(strainTheoryShapeTitleVar(i)).value = strainTheoryShapeTitle(i)
        
        '�����ͼ���ˣ�����
        If StrainTheoryShapeObjCounts = strainNWCs Then
            StrainTheoryShapeObjArray(i).CopyPicture
            wordApp.ActiveDocument.Bookmarks(strainTheoryShapeBookmarks(i)).Range.Paste
        End If
    Next
    
    wordApp.ActiveDocument.Fields.Update         '������
    
    DelCalcSpecifiedBookmarks wordApp.ActiveDocument
    UnlinkCalcSpecifiedDocVarFields wordApp.ActiveDocument

    '�Է���һ���½���ǩ���ر�����MS Word�У�
    wordApp.ActiveDocument.Application.CaptionLabels.Add name:="ͼ"
    wordApp.ActiveDocument.Application.CaptionLabels.Add name:="��"
    
    'TODO:���ݱ���ģ���޸Ĳ���
'    Dim tableOffset As Integer
'    Dim chartOffset As Integer
'    tableOffset = 1
'    chartOffset = 5
    
    '������ע
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
        MsgBox "�����鵼��ʧ�ܣ�"
    End If
End Sub

'���ȼ��㣬������Word����
Public Sub AutoReport()

    'Kill ThisWorkbook.Path & "\AutoReportSource.docx"
    GenAutoReportTemplate                        '�ȳ�ʼ��ģ��
    
    Dim reportTemplateFileName As String
    reportTemplateFileName = "�Զ�����ģ��.docx"       '"����ģ��.docx"

    Dim reportFileName As String
    reportFileName = "�Զ����ɵı���.docx"
    
    
    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    
    Dim i, j As Integer
    Dim resultFlag As Boolean                    'True��ʾ�����ɹ�
    resultFlag = False
    
    Dim dispResultVar(1 To MAX_NWC) As String    '��word����ӦDocVariable��Ӧ���������飩
    For i = 1 To MAX_NWC
        dispResultVar(i) = Replace("dispResult" & Str(i), " ", "")
    Next
    
    Dim dispSummaryVar(1 To MAX_NWC) As String   '��word����ӦDocVariable��Ӧ���������飩
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

    Dim tableDataStartRow As Integer             '���������ʼ��
    
    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\" & reportTemplateFileName, ThisWorkbook.Path & "\AutoReportSource.docx"
    
    wordApp.Documents.Open ThisWorkbook.Path & "\AutoReportSource.docx"

    wordApp.Visible = False
    
    tableDataStartRow = 3
    For i = 1 To NWCs                            'i��λ����
        wordApp.ActiveDocument.Variables(dispResultVar(i)).value = dispResult(i)
        If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOff Then 'ע��ֻ��ʹ��xlOn������ʹ��True��������ֵΪ1��
            dispSummary(i) = Replace(dispSummary(i), "���Ӷ�ʵ��ֵ�����ۼ���ֵ�Ĺ�ϵ�������" & dispGraphCrossRef(i), "")
        End If
        wordApp.ActiveDocument.Variables(dispSummaryVar(i)).value = dispSummary(i)
        wordApp.ActiveDocument.Variables(dispTbTitleVar(i)).value = dispTbTitle(i)
        
        
        '������
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(dispTblBookmarks(i)).Range, NumRows:=DispUbound(i) + tableDataStartRow - 1, NumColumns:=7) 'NumRows+x��ʾ��ͷ
        
        SetDispTableWidth tbl, DispUbound(i) + tableDataStartRow - 1    '���ø��п��
               
        tbl.Cell(1, 1).Merge tbl.Cell(2, 1)
        tbl.Cell(1, 2).Merge tbl.Cell(1, 4)
        tbl.Cell(1, 3).Merge tbl.Cell(2, 5)
        tbl.Cell(1, 4).Merge tbl.Cell(2, 6)
        tbl.Cell(1, 5).Merge tbl.Cell(2, 7)
        
        tbl.Cell(1, 1).Range.InsertAfter "����"
        tbl.Cell(1, 2).Range.InsertAfter "ʵ��ֵ(mm)"
        tbl.Cell(2, 2).Range.InsertAfter "�ܱ���"
        tbl.Cell(2, 3).Range.InsertAfter "���Ա���"
        tbl.Cell(2, 4).Range.InsertAfter "�������"
        tbl.Cell(1, 3).Range.InsertAfter "��������ֵ(mm)"
        tbl.Cell(1, 4).Range.InsertAfter "У��ϵ��"
        tbl.Cell(1, 5).Range.InsertAfter "��Բ������"
        
        tbl.Cell(1, 1).Range.Font.Bold = True
        tbl.Cell(1, 2).Range.Font.Bold = True
        tbl.Cell(2, 2).Range.Font.Bold = True
        tbl.Cell(2, 3).Range.Font.Bold = True
        tbl.Cell(2, 4).Range.Font.Bold = True
        tbl.Cell(1, 3).Range.Font.Bold = True
        tbl.Cell(1, 4).Range.Font.Bold = True
        tbl.Cell(1, 5).Range.Font.Bold = True
        
        
        For j = 1 To DispUbound(i)               'j��λ���
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
        
        '�Ƿ񵼳�ͼ��
        If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then 'ע��ֻ��ʹ��xlOn������ʹ��True��������ֵΪ1��
             wordApp.ActiveDocument.Variables(dispChartTitleVar(i)).value = dispChartTitle(i)
            DispChartObjArray(i).Copy
            wordApp.ActiveDocument.Bookmarks(dispChartBookmarks(i)).Range.Paste
        End If
        
    Next i
    tbl.Rows.Alignment = wdAlignRowCenter
    
    tableDataStartRow = 3
    For i = 1 To strainNWCs                      'i��λ����
      
        wordApp.ActiveDocument.Variables(strainResultVar(i)).value = strainResult(i)
        If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOff Then 'ע��ֻ��ʹ��xlOn������ʹ��True��������ֵΪ1��
            strainSummary(i) = Replace(strainSummary(i), "��Ӧ��ʵ��ֵ�����ۼ���ֵ�Ĺ�ϵ�������" & strainGraphCrossRef(i), "")
        End If
        wordApp.ActiveDocument.Variables(strainSummaryVar(i)).value = strainSummary(i)
        wordApp.ActiveDocument.Variables(strainTbTitleVar(i)).value = strainTbTitle(i)
        
        '������
        
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(strainTblBookmarks(i)).Range, NumRows:=StrainUbound(i) + tableDataStartRow - 1, NumColumns:=7) 'NumRows+1��ʾ��ͷ
        
        SetStrainTableWidth tbl, StrainUbound(i) + tableDataStartRow - 1
        tbl.Cell(1, 1).Merge tbl.Cell(2, 1)
        tbl.Cell(1, 2).Merge tbl.Cell(1, 4)
        tbl.Cell(1, 3).Merge tbl.Cell(2, 5)
        tbl.Cell(1, 4).Merge tbl.Cell(2, 6)
        tbl.Cell(1, 5).Merge tbl.Cell(2, 7)
        
        tbl.Cell(1, 1).Range.InsertAfter "����"
        tbl.Cell(1, 2).Range.InsertAfter "ʵ��ֵ(�̦�)"
        tbl.Cell(2, 2).Range.InsertAfter "��Ӧ��"
        tbl.Cell(2, 3).Range.InsertAfter "����Ӧ��"
        tbl.Cell(2, 4).Range.InsertAfter "����Ӧ��"
        tbl.Cell(1, 3).Range.InsertAfter "��������ֵ(�̦�)"
        tbl.Cell(1, 4).Range.InsertAfter "У��ϵ��"
        tbl.Cell(1, 5).Range.InsertAfter "��Բ���Ӧ��"
        
        tbl.Cell(1, 1).Range.Font.Bold = True
        tbl.Cell(1, 2).Range.Font.Bold = True
        tbl.Cell(2, 2).Range.Font.Bold = True
        tbl.Cell(2, 3).Range.Font.Bold = True
        tbl.Cell(2, 4).Range.Font.Bold = True
        tbl.Cell(1, 3).Range.Font.Bold = True
        tbl.Cell(1, 4).Range.Font.Bold = True
        tbl.Cell(1, 5).Range.Font.Bold = True
        
        For j = 1 To StrainUbound(i)             'j��λ���
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
        
        '�Ƿ񵼳�ͼ��
        If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then 'ע��ֻ��ʹ��xlOn������ʹ��True��������ֵΪ1��
            wordApp.ActiveDocument.Variables(strainChartTitleVar(i)).value = strainChartTitle(i)
            StrainChartObjArray(i).Copy
            wordApp.ActiveDocument.Bookmarks(strainChartBookmarks(i)).Range.Paste
        End If
        
    Next i
   'wordApp.ActiveDocument.Tables(2).Rows.Alignment = wdAlignRowCenter
    
    wordApp.ActiveDocument.Fields.Update         '������

    'DelAllBookmarks wordApp.ActiveDocument       'ɾ��������ǩ
    DelSpecifiedBookmarks wordApp.ActiveDocument
    UnlinkSpecifiedDocVarFields wordApp.ActiveDocument
    'UnlinkAllFields wordApp.ActiveDocument       '�������������

    '�Է���һ���½���ǩ���ر�����MS Word�У�
    wordApp.ActiveDocument.Application.CaptionLabels.Add name:="ͼ"
    wordApp.ActiveDocument.Application.CaptionLabels.Add name:="��"
    
    'TODO:���ݱ���ģ���޸Ĳ���
    Dim tableOffset As Integer
    Dim chartOffset As Integer
    tableOffset = 1
    chartOffset = 5
    
    '������ע
    SetAutoReportCaptions wordApp
    '���ý�������
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
        MsgBox "���浼��ʧ�ܣ�"
    End If
    
End Sub

'�򿪺�������word����
Public Sub OpenReport()
    Dim wordApp As Word.Application
    Set wordApp = CreateObject("Word.Application")
    wordApp.Documents.Open fileName:=ThisWorkbook.Path & "\" & AutoReportFileName, ReadOnly:=False
    wordApp.Visible = True
    wordApp.Activate
End Sub

'�򿪺�����������鱨��
Public Sub OpenCalcReport()
    Dim wordApp As Word.Application
    Set wordApp = CreateObject("Word.Application")
    wordApp.Documents.Open fileName:=ThisWorkbook.Path & "\" & AutoCalcReportFileName, ReadOnly:=False
    wordApp.Visible = True
    wordApp.Activate
End Sub

'���Ա������ɲ���
Public Sub testTable()
    Dim i, j, k As Integer

    Dim templateFileName As String               'ģ���ļ�
    templateFileName = "��������ģ��.docx"             '���������������������Ҫ����Դ
    
    Dim reportTemplateFileName As String
    reportTemplateFileName = "������������ģ��.docx"     '"����ģ��.docx"
    
    Dim reportFileName As String
    reportFileName = "�Զ����ɵ�������������test.docx"
    
    Dim tempFileName As String                   '��ʱ�ļ�������
    tempFileName = "temp.docx"
    
    Dim tempTemplateFileName As String           '��ʱģ���ļ�������
    tempFileName = "temp.docx"
    
    Dim wordApp As Word.Application
    
    Dim tempDoc As Word.Document
    Dim templateDoc As New Word.Document
    
    Dim tbl As Word.Table                        'As Object    'Variant/Object/Range
    Dim tempTable As Word.Table
    
    Dim r As Word.Range
    
    
    Dim resultFlag As Boolean                    'True��ʾ�����ɹ�
    resultFlag = False
    
    Dim tbRowStart As Integer

    
    Dim tableDataStartRow As Integer             '���������ʼ��
    
    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\" & reportTemplateFileName, ThisWorkbook.Path & "\" & tempFileName
    FileCopy ThisWorkbook.Path & "\��������ģ��bak.docx", ThisWorkbook.Path & "\" & templateFileName
    
    Set tempDoc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & tempFileName)
    Set templateDoc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & templateFileName)

    wordApp.Visible = True

    
    tbRowStart = 2
    

    
    Set tbl = templateDoc.Tables(1)              '��ʱ��֧�ַ�5������
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

     
    tempDoc.Fields.Update                        '������
    
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
        MsgBox "���浼����ɣ�"
    Else
        MsgBox "���浼��ʧ�ܣ�"
    End If
    
End Sub

'TODO����֧��10�����
'���ȼ��㣬�����ɱ��棨����������
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
    
    
    Dim templateFileName As String               'ģ���ļ�
    templateFileName = "��������ģ��.docx"             '���������������������Ҫ����Դ
    
    Dim reportTemplateFileName As String
    reportTemplateFileName = "������������ģ��.docx"     '"����ģ��.docx"
    
    Dim reportFileName As String
    reportFileName = "�Զ����ɵ�������������.docx"
    
    Dim tempFileName As String                   '��ʱ�ļ�������
    tempFileName = "temp.docx"
    
    Dim tempTemplateFileName As String           '��ʱģ���ļ�������
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
    
    
    Dim resultFlag As Boolean                    'True��ʾ�����ɹ�
    resultFlag = False
    
    Dim tbRowStart As Integer
    Dim tableDataStartRow As Integer             '���������ʼ��
    
    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\" & reportTemplateFileName, ThisWorkbook.Path & "\" & tempFileName
    FileCopy ThisWorkbook.Path & "\��������ģ��bak.docx", ThisWorkbook.Path & "\" & templateFileName
    
    Set tempDoc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & tempFileName)
    Set templateDoc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & templateFileName)

    wordApp.Visible = False
    
    tbRowStart = 2
    
    For i = 1 To nThrust
    
        '����������
        thrustResult(i) = "����������������������£����˱����ȶ����������Ա���ֵΪ" & Format(ThrustElasticDisp(i), "Fixed") & "mm��" _
                                                                                                  & "�������" & Format(ThrustRemainDisp(i), "Fixed") & "mm��" _
                                                                                                  & "��Բ������Ϊ" & Format(ThrustRefRemainDisp(i), "Percent") & "��"
        
        tempDoc.Variables(thrustResultVar(i)).value = thrustResult(i)

        '���������
        tableTitle(i) = "��x-x " & CStr(i) & "#�����μ�������ܱ�"
        tempDoc.Variables(tableTitleVar(i)).value = tableTitle(i)
    
    
        Set tbl = templateDoc.Tables(i)          '��ʱ��֧�ַ�5������
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
     
    tempDoc.Fields.Update                        '������

    DelAllBookmarks tempDoc                      'ɾ��������ǩ
    UnlinkAllFields tempDoc                      '�������������
    
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
        MsgBox "���浼����ɣ�"
    Else
        MsgBox "���浼��ʧ�ܣ�"
    End If
    
End Sub

'����table�����ֶ��뷽ʽ
Private Sub SetTableAlignment(ByRef tbl As Table)
    On Error Resume Next
    tbl.Select
    tbl.Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    tbl.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    tbl.Range.Rows.Alignment = wdAlignRowCenter
End Sub

'����table������Ԫ����
Private Sub SetDispTableWidth(ByRef tbl As Table, ByVal NumRows As Integer)
    On Error Resume Next
    Dim NumColumns As Integer
    NumColumns = 7
    Dim WidthArray(1 To 7) As Single    '���п��
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

'����λ��ԭʼ���ݴ���������Ԫ����
Private Sub SetStrainRawTableWidth(ByRef tbl As Table, ByVal NumRows As Integer)
    On Error Resume Next
    Dim NumColumns As Integer
    NumColumns = 14
    Dim WidthArray(1 To 14) As Single    '���п��
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

'����table������Ԫ����
Private Sub SetStrainTableWidth(ByRef tbl As Table, ByVal NumRows As Integer)
    On Error Resume Next
    Dim NumColumns As Integer
    NumColumns = 7
    Dim WidthArray(1 To 7) As Single    '���п��
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

'���ü�����Ӧ���������Ԫ����
Private Sub SetCalcStrainTableWidth(ByRef tbl As Table, ByVal NumRows As Integer)
    On Error Resume Next
    Dim NumColumns As Integer
    NumColumns = 8
    Dim WidthArray(1 To 8) As Single    '���п��
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


'����table�ı߽���
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

'���ָ������ǩ
Private Sub DelSpecifiedBookmarks(ByRef doc As Word.Document)

    DelArrayBookmarks doc, dispTblBookmarks
    DelArrayBookmarks doc, strainTblBookmarks
    
    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then 'ע��ֻ��ʹ��xlOn������ʹ��True��������ֵΪ1��
        DelArrayBookmarks doc, dispChartBookmarks
        DelArrayBookmarks doc, strainChartBookmarks
    End If
End Sub

'���������ָ������ǩ
Private Sub DelCalcSpecifiedBookmarks(ByRef doc As Word.Document)

    DelArrayBookmarks doc, dispRawTblBookmarks
    DelArrayBookmarks doc, dispTblBookmarks
    DelArrayBookmarks doc, dispTheoryShapeBookmarks
    DelArrayBookmarks doc, strainTblBookmarks
    DelArrayBookmarks doc, strainRawTableBookmarks
    DelArrayBookmarks doc, strainTheoryShapeBookmarks
    
    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox_Calc") = xlOn Then 'ע��ֻ��ʹ��xlOn������ʹ��True��������ֵΪ1��
        DelArrayBookmarks doc, dispChartBookmarks
        DelArrayBookmarks doc, strainChartBookmarks
    End If
End Sub

'����һ���ַ�������������Ԫ����������ǩɾ��
Private Sub DelArrayBookmarks(ByRef doc As Word.Document, ByRef arr() As String)
    On Error Resume Next
    Dim i As Integer
    For i = 1 To UBound(arr)
        doc.Bookmarks(arr(i)).Delete
    Next i
End Sub

'ɾ��������ǩ
Private Sub DelAllBookmarks(ByRef doc As Word.Document)
    On Error Resume Next
    Dim bk As Bookmark
    For Each bk In doc.Bookmarks
        bk.Delete
    Next
End Sub

'����һ���ַ�������������Ԫ���������ĵ������������
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

'���ָ�������ĵ�����������
Private Sub UnlinkSpecifiedDocVarFields(ByRef doc As Word.Document)

    UnlinkArrayFields doc, dispSummary
    UnlinkArrayFields doc, dispResult
    UnlinkArrayFields doc, dispTbTitle

    UnlinkArrayFields doc, strainSummary
    UnlinkArrayFields doc, strainResult
    UnlinkArrayFields doc, strainTbTitle
    
    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox") = xlOn Then 'ע��ֻ��ʹ��xlOn������ʹ��True��������ֵΪ1��
        UnlinkArrayFields doc, dispChartTitle
        UnlinkArrayFields doc, strainChartTitle
    End If
End Sub

'���ָ�������ĵ�����������
Private Sub UnlinkCalcSpecifiedDocVarFields(ByRef doc As Word.Document)

    UnlinkArrayFields doc, dispRawTbTitle
    UnlinkArrayFields doc, dispTbTitle
    UnlinkArrayFields doc, dispTheoryShapeTitle
    UnlinkArrayFields doc, strainRawTbTitle
    UnlinkArrayFields doc, strainTbTitle
    UnlinkArrayFields doc, strainTheoryShapeTitle
    
    If ActiveSheet.CheckBoxes("ExportRelationChartCheckBox_Calc") = xlOn Then 'ע��ֻ��ʹ��xlOn������ʹ��True��������ֵΪ1��
        UnlinkArrayFields doc, dispChartTitle
        UnlinkArrayFields doc, strainChartTitle
    End If
End Sub

'�������������
Private Sub UnlinkAllFields(ByRef doc As Word.Document)
    On Error Resume Next
    Dim f As Field
    For Each f In doc.Fields
        f.Unlink
    Next

End Sub


