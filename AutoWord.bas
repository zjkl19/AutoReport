Attribute VB_Name = "AutoWord"
'���Ӵ���
'���Կ��У�����ǩ���������
'wordApp.ActiveDocument.Bookmarks("dispSummary1").Range.InsertAfter "111"
'���Բ��������
'wordApp.ActiveDocument.Tables.Add wordApp.ActiveDocument.Bookmarks("dispSummary1").Range, NumRows:=14 + 1, NumColumns:=7


'���ȼ��㣬������Word����
Sub AutoWordDisp()
    'On Error GoTo CloseWord
    
    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    
    Dim i, j As Integer
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
    
    Dim strainSummary(1 To MAX_NWC) As String    '��word����ӦDocVariable��Ӧ
    Dim strainSummaryVar(1 To MAX_NWC) As String    '��word����ӦDocVariable��Ӧ
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
    
        '�������
        dispSummary(i) = "����" & CStr(nPN(GlobalWC(i) - 1)) & "���Խ������Ӷȼ���������x-x��ͼx-x������������������������Ӷ�У��ϵ����" & Format(StatPara(i, 1), "Fixed") & "��" & Format(StatPara(i, 2), "Fixed") & "֮�䣬" _
        & "���㡶��·���������������������̡��й涨��У��ϵ��С��1.0��Ҫ�����������������Բ������Ϊ" & Format(StatPara(i, 3), "Percent") & "��" _
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
    
        '�������
        strainSummary(i) = "����" & CStr(nPN(StrainGlobalWC(i) - 1)) & "���Խ�����Ӧ�����������x-x��ͼx-x�����������������������Ӧ��У��ϵ����" _
        & Format(StrainStatPara(i, 1), "Fixed") & "��" & Format(StrainStatPara(i, 2), "Fixed") & "֮�䣬" _
        & "���㡶��·���������������������̡��涨��У��ϵ��С��1.0��Ҫ�����⹹���������Բ���Ӧ��Ϊ" & Format(StrainStatPara(i, 3), "Percent") _
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
            tbl.Cell(1 + j, 2).Range.InsertAfter Format(TotalStrain(i, j), "Fixed")
            tbl.Cell(1 + j, 3).Range.InsertAfter Format(ElasticStrain(i, j), "Fixed")
            tbl.Cell(1 + j, 4).Range.InsertAfter Format(RemainStrain(i, j), "Fixed")
            tbl.Cell(1 + j, 5).Range.InsertAfter Format(TheoryStrain(i, j), "Fixed")
            tbl.Cell(1 + j, 6).Range.InsertAfter Format(StrainCheckoutCoff(i, j), "Fixed")
            tbl.Cell(1 + j, 7).Range.InsertAfter Format(RefRemainStrain(i, j), "Percent")
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
