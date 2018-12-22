Attribute VB_Name = "AutoWord"
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

'�Զ����ɼ�����
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

    For i = 1 To nWCs    'i��λ����
        
        '���������
        dispRawTbTitle(i) = "��x-x ����" & CStr(nPN(GlobalWC(i) - 1)) & "�Ӷ�ԭʼ���ݴ����"
        wordApp.ActiveDocument.Variables(dispRawTbTitleVar(i)).value = dispRawTbTitle(i)
        
        '������
        
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(dispRawTblBookmarks(i)).Range, NumRows:=DispUbound(i) + 1, NumColumns:=7)    'NumRows+1��ʾ��ͷ
        tbl.Cell(1, 1).Range.InsertAfter "����"
        tbl.Cell(1, 2).Range.InsertAfter "��ʼ����"
        tbl.Cell(1, 3).Range.InsertAfter "����"
        tbl.Cell(1, 4).Range.InsertAfter "����"
        tbl.Cell(1, 5).Range.InsertAfter "���Ӷ�"
        tbl.Cell(1, 6).Range.InsertAfter "�����Ӷ�"
        tbl.Cell(1, 7).Range.InsertAfter "�������"
        

        For j = 1 To DispUbound(i)    'j��λ���
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
    
    For i = 1 To nWCs    'i��λ����
        
        '���������
        dispTbTitle(i) = "��x-x ����" & CStr(nPN(GlobalWC(i) - 1)) & "�Ӷȼ�������ܱ�"
        wordApp.ActiveDocument.Variables(dispTbTitleVar(i)).value = dispTbTitle(i)
        
        '������
        
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(dispTblBookmarks(i)).Range, NumRows:=DispUbound(i) + 1, NumColumns:=7)    'NumRows+1��ʾ��ͷ
        tbl.Cell(1, 1).Range.InsertAfter "����"
        tbl.Cell(1, 2).Range.InsertAfter "�ܱ���"
        tbl.Cell(1, 3).Range.InsertAfter "���Ա���"
        tbl.Cell(1, 4).Range.InsertAfter "�������"
        tbl.Cell(1, 5).Range.InsertAfter "��������ֵ(mm)"
        tbl.Cell(1, 6).Range.InsertAfter "У��ϵ��"
        tbl.Cell(1, 7).Range.InsertAfter "��Բ������(%)"
        

        For j = 1 To DispUbound(i)    'j��λ���
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

    For i = 1 To StrainNWCs    'i��λ����
  
        '���������
        strainRawTableTitle(i) = "��x-x ����" & CStr(nPN(StrainGlobalWC(i) - 1)) & "Ӧ��ԭʼ���ݴ����"
        wordApp.ActiveDocument.Variables(strainRawTableTitleVar(i)).value = strainRawTableTitle(i)
        
        '������
        
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(strainRawTableBookmarks(i)).Range, NumRows:=StrainUbound(i) + 1, NumColumns:=14)    'NumRows+1��ʾ��ͷ
        tbl.Cell(1, 1).Range.InsertAfter "����"
        tbl.Cell(1, 2).Range.InsertAfter "ģ��R"
        tbl.Cell(1, 3).Range.InsertAfter "�¶�T"
        tbl.Cell(1, 4).Range.InsertAfter "ģ��R"
        tbl.Cell(1, 5).Range.InsertAfter "�¶�T"
        tbl.Cell(1, 6).Range.InsertAfter "ģ��R"
        tbl.Cell(1, 7).Range.InsertAfter "�¶�T"
        tbl.Cell(1, 8).Range.InsertAfter "��R"
        tbl.Cell(1, 9).Range.InsertAfter "��T"
        tbl.Cell(1, 10).Range.InsertAfter "��R"
        tbl.Cell(1, 11).Range.InsertAfter "��T"
        tbl.Cell(1, 12).Range.InsertAfter "��Ӧ��"
        tbl.Cell(1, 13).Range.InsertAfter "����Ӧ��"
        tbl.Cell(1, 14).Range.InsertAfter "����Ӧ��"
        
        For j = 1 To StrainUbound(i)    'j��λ���
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

    For i = 1 To StrainNWCs    'i��λ����
  
  
        '���������
        strainTableTitle(i) = "��x-x ����" & CStr(nPN(StrainGlobalWC(i) - 1)) & "Ӧ���������ܱ�"
        wordApp.ActiveDocument.Variables(strainTableTitleVar(i)).value = strainTableTitle(i)
        
        '������
        
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(strainTableBookmarks(i)).Range, NumRows:=StrainUbound(i) + 1, NumColumns:=8)    'NumRows+1��ʾ��ͷ
        tbl.Cell(1, 1).Range.InsertAfter "����"
        tbl.Cell(1, 2).Range.InsertAfter "��Ӧ��"
        tbl.Cell(1, 3).Range.InsertAfter "����Ӧ��"
        tbl.Cell(1, 4).Range.InsertAfter "����Ӧ��"
        tbl.Cell(1, 5).Range.InsertAfter "����Ӧ������ֵ��MPa��"
        tbl.Cell(1, 6).Range.InsertAfter "��������ֵ(�̦�)"
        tbl.Cell(1, 7).Range.InsertAfter "У��ϵ��"
        tbl.Cell(1, 8).Range.InsertAfter "��Բ���Ӧ��(%)"
        
        For j = 1 To StrainUbound(i)    'j��λ���
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
    
    
    wordApp.ActiveDocument.Fields.Update    '������

    wordApp.Documents.Save
    
    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\AutoCalcReportResult.docx"
    
    Set wordApp = Nothing
    Set tbl = Nothing
End Sub

'���ȼ��㣬������Word����
Public Sub AutoReport()
    'On Error GoTo CloseWord
    
    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    
    Dim i, j As Integer
    
    Dim dispResult(1 To MAX_NWC) As String    '��word����ӦDocVariable��Ӧ
    Dim dispResultVar(1 To MAX_NWC) As String    '��word����ӦDocVariable��Ӧ���������飩
    For i = 1 To MAX_NWC
        dispResultVar(i) = Replace("dispResult" & GlobalWC(i), " ", "")
    Next
    
    Dim dispSummary(1 To MAX_NWC) As String    '��word����ӦDocVariable��Ӧ
    Dim dispSummaryVar(1 To MAX_NWC) As String    '��word����ӦDocVariable��Ӧ���������飩
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

    For i = 1 To nWCs    'i��λ����
        '����������
        dispResult(i) = "(" & Str(i) & ")�ڹ���" & CStr(nPN(GlobalWC(i) - 1)) & "���������£��������ʵ�ⵯ���Ӷ�ֵΪ" & Format(StatPara(i, MaxElasticDeform_Index), "Fixed") & "mm��" _
        & "ʵ����ƽ�����Ӷ�ֵ��С������ֵ��У��ϵ����" & Format(StatPara(i, MinCheckoutCoff_Index), "Fixed") & "��" & Format(StatPara(i, MaxCheckoutCoff_Index), "Fixed") & "֮�䣻" _
        & "��Բ��������" & Format(StatPara(i, MinRefRemainDeform_Index), "Percent") & "��" & Format(StatPara(i, MaxRefRemainDeform_Index), "Percent") & "֮�䡣"
        
        wordApp.ActiveDocument.Variables(dispResultVar(i)).value = dispResult(i)
    
        '�������
        dispSummary(i) = "����" & CStr(nPN(GlobalWC(i) - 1)) & "���Խ������Ӷȼ���������x-x��ͼx-x������������������������Ӷ�У��ϵ����" & Format(StatPara(i, MinCheckoutCoff_Index), "Fixed") & "��" & Format(StatPara(i, MaxCheckoutCoff_Index), "Fixed") & "֮�䣬" _
        & "���㡶��·���������������������̡��й涨��У��ϵ��С��1.0��Ҫ�����������������Բ������Ϊ" & Format(StatPara(i, MaxRefRemainDeform_Index), "Percent") & "��" _
        & "���㡶��·���������������������̡��й涨�Ĳ��������ֵҪ��(��ֵ20%)���ָ�״�����á�"
               
        wordApp.ActiveDocument.Variables(dispSummaryVar(i)).value = dispSummary(i)
        
        '���������
        dispTbTitle(i) = "��x-x ����" & CStr(nPN(GlobalWC(i) - 1)) & "�Ӷȼ�������ܱ�"
        wordApp.ActiveDocument.Variables(dispTbTitleVar(i)).value = dispTbTitle(i)
        
        '������
        
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(dispTblBookmarks(i)).Range, NumRows:=DispUbound(i) + 1, NumColumns:=7)    'NumRows+1��ʾ��ͷ
        tbl.Cell(1, 1).Range.InsertAfter "����"
        tbl.Cell(1, 2).Range.InsertAfter "�ܱ���"
        tbl.Cell(1, 3).Range.InsertAfter "���Ա���"
        tbl.Cell(1, 4).Range.InsertAfter "�������"
        tbl.Cell(1, 5).Range.InsertAfter "��������ֵ(mm)"
        tbl.Cell(1, 6).Range.InsertAfter "У��ϵ��"
        tbl.Cell(1, 7).Range.InsertAfter "��Բ������(%)"
        

        For j = 1 To DispUbound(i)    'j��λ���
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
    
    For i = 1 To StrainNWCs    'i��λ����
        '����������
        strainResult(i) = "(" & Str(i) & ")�ڹ���" & CStr(nPN(StrainGlobalWC(i) - 1)) & "���������£��������������Ӧ��Ϊ" & Format(StrainStatPara(i, MaxElasticDeform_Index), "Fixed") & "�̦ţ�" _
        & "ʵ����ƽ���Ļ�����Ӧ��ֵ��С������ֵ��У��ϵ����" & Format(StrainStatPara(i, MinCheckoutCoff_Index), "Fixed") & "��" & Format(StrainStatPara(i, MaxCheckoutCoff_Index), "Fixed") & "֮�䣻" _
        & "��Բ���Ӧ����" & Format(StrainStatPara(i, MinRefRemainDeform_Index), "Percent") & "��" & Format(StrainStatPara(i, MaxRefRemainDeform_Index), "Percent") & "֮�䡣"
        
        wordApp.ActiveDocument.Variables(strainResultVar(i)).value = strainResult(i)
        '�������
        strainSummary(i) = "����" & CStr(nPN(StrainGlobalWC(i) - 1)) & "���Խ�����Ӧ�����������x-x��ͼx-x�����������������������Ӧ��У��ϵ����" _
        & Format(StrainStatPara(i, MinCheckoutCoff_Index), "Fixed") & "��" & Format(StrainStatPara(i, MaxCheckoutCoff_Index), "Fixed") & "֮�䣬" _
        & "���㡶��·���������������������̡��涨��У��ϵ��С��1.0��Ҫ�����⹹���������Բ���Ӧ��Ϊ" & Format(StrainStatPara(i, MaxRefRemainDeform_Index), "Percent") _
        & "�����㡶��·���������������������̡��й涨�Ĳ���Ӧ����ֵҪ��(��ֵ20%)���ָ�״�����á�"
        
        wordApp.ActiveDocument.Variables(strainSummaryVar(i)).value = strainSummary(i)
        
        '���������
        strainTbTitle(i) = "��x-x ����" & CStr(nPN(StrainGlobalWC(i) - 1)) & "Ӧ���������ܱ�"
        wordApp.ActiveDocument.Variables(strainTbTitleVar(i)).value = strainTbTitle(i)
        
        '������
        
        Set tbl = wordApp.ActiveDocument.Tables.Add(wordApp.ActiveDocument.Bookmarks(strainTblBookmarks(i)).Range, NumRows:=StrainUbound(i) + 1, NumColumns:=7)    'NumRows+1��ʾ��ͷ
        tbl.Cell(1, 1).Range.InsertAfter "����"
        tbl.Cell(1, 2).Range.InsertAfter "��Ӧ��"
        tbl.Cell(1, 3).Range.InsertAfter "����Ӧ��"
        tbl.Cell(1, 4).Range.InsertAfter "����Ӧ��"
        tbl.Cell(1, 5).Range.InsertAfter "��������ֵ(�̦�)"
        tbl.Cell(1, 6).Range.InsertAfter "У��ϵ��"
        tbl.Cell(1, 7).Range.InsertAfter "��Բ���Ӧ��(%)"
        
        For j = 1 To StrainUbound(i)    'j��λ���
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
    
    wordApp.ActiveDocument.Fields.Update    '������

    wordApp.Documents.Save
    
    wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\AutoReportResult.docx"
CloseWord:
    wordApp.Documents.Close
    wordApp.Quit
    
    Set wordApp = Nothing
    Set tbl = Nothing
    
End Sub
'����table�ı߽���
Private Sub SetTableBorder(ByRef tbl As Table)
    With tbl
        With .Borders
            .InsideLineStyle = wdLineStyleSingle
            .OutsideLineStyle = wdLineStyleSingle
        End With
    End With
End Sub
