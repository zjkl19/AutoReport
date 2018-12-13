Attribute VB_Name = "AutoDisp"
Option Explicit
Const First_Row As Integer = 8     '��ʼ��������
Const WC_Col As Integer = 1     '������������
Const MAX_NWC As Integer = 10     '��󹤿���

Const Node_Name_Col As Integer = 2  '�����������
Const TheoryDisp_Col As Integer = 10  '����λ��������
Dim TotalDispCol As Integer    '�ܱ���������
Dim DeltaCol As Integer   '����������
Dim RemainDispCol As Integer    '�������������
Dim ElasticCol As Integer    '���Ա���������
Dim CheckoutCoffCol As Integer    'У��ϵ��������
Dim RefRemainDispCol As Integer    '��Բ������������

Dim nWCs As Integer    '������
Dim nPN    '����������Ӧ��������
'nPN = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ")

Dim TotalDisp(1 To MAX_NWC, 1 To 100)   'TotalDisp(i,j)��ʾ��i����������j������ܱ���
Dim NodeName(1 To MAX_NWC, 1 To 100) As String  '���������������
Dim Delta(1 To MAX_NWC, 1 To 100)
Dim RemainDisp(1 To MAX_NWC, 1 To 100)
Dim ElasticDisp(1 To MAX_NWC, 1 To 100)
Dim TheoryDisp(1 To MAX_NWC, 1 To 100)
Dim CheckoutCoff(1 To MAX_NWC, 1 To 100)
Dim RefRemainDisp(1 To MAX_NWC, 1 To 100)
Dim DispUbound(1 To MAX_NWC) As Integer    'ÿ�������Ͻ磨�½�Ϊ1��

Dim StatPara(1 To MAX_NWC, 1 To 3)  'ͳ�Ʋ���,��СУ��ϵ�������У��ϵ���������Բ���Ӧ��
'StatPara(i,1~3)�ֱ��ʾ��i��������СУ��ϵ�������У��ϵ���������Բ���Ӧ��

Dim t

'''��ʼ��ȫ�ֱ���
Public Sub InitVar()

    nPN = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ")
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
    
    Dim rowCurr As Integer    '��ָ��
    
    Dim i, j As Integer
    rowCurr = First_Row
    For i = 1 To nWCs
        For j = 1 To DispUbound(i)
        
            NodeName(i, j) = Cells(rowCurr, Node_Name_Col)
            TheoryDisp(i, j) = Cells(rowCurr, TheoryDisp_Col)
            
            TotalDisp(i, j) = Cells(rowCurr, TotalDispCol - 1) - Cells(rowCurr, TotalDispCol - 2)    '�ܱ���
            Cells(rowCurr, TotalDispCol) = TotalDisp(i, j)
            
            Cells(rowCurr, DeltaCol) = Cells(rowCurr, DeltaCol - 1) - Cells(rowCurr, DeltaCol - 3)   '����
            '�������洢
            
            '�㷨��ж�������ض�����ֵ>=0��ȡж�������ض�����ֵ������ȡ0
            RemainDisp(i, j) = IIf(Cells(rowCurr, RemainDispCol - 2) - Cells(rowCurr, RemainDispCol - 5) >= 0, Cells(rowCurr, RemainDispCol - 2) - Cells(rowCurr, RemainDispCol - 5), 0)
            Cells(rowCurr, RemainDispCol) = RemainDisp(i, j)    '�������
            
            ElasticDisp(i, j) = Cells(rowCurr, ElasticCol - 4) - Cells(rowCurr, ElasticCol - 1)
            Cells(rowCurr, ElasticCol) = ElasticDisp(i, j)    '���Ա���
             
            CheckoutCoff(i, j) = Cells(rowCurr, CheckoutCoffCol - 2) / Cells(rowCurr, CheckoutCoffCol - 1)
            Cells(rowCurr, CheckoutCoffCol) = CheckoutCoff(i, j)    'У��ϵ��
             
            RefRemainDisp(i, j) = Cells(rowCurr, RefRemainDispCol - 4) / Cells(rowCurr, RefRemainDispCol - 7)
            Cells(rowCurr, RefRemainDispCol) = RefRemainDisp(i, j)    '��Բ������
            
            rowCurr = rowCurr + 1
        Next
    Next
    
    '�������������С/��У��ϵ���������Բ������
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
        
        '����д��Excel
        Cells(3, 2 * i) = Format(StatPara(i, 1), "Fixed"): Cells(4, 2 * i) = Format(StatPara(i, 2), "Fixed"): Cells(5, 2 * i) = Format(StatPara(i, 3), "Percent")
    Next
'    While Cells(rowCurr, 1) <> ""
'        Cells(rowCurr, TotalDispCol) = Cells(rowCurr, TotalDispCol - 1) - Cells(rowCurr, TotalDispCol - 2)
'        Cells(rowCurr, DeltaCol) = Cells(rowCurr, DeltaCol - 1) - Cells(rowCurr, DeltaCol - 3)
'
'        �㷨��ж�������ض�����ֵ>=0��ȡж�������ض�����ֵ������ȡ0
'        Cells(rowCurr, RemainDispCol) = IIf(Cells(rowCurr, RemainDispCol - 2) - Cells(rowCurr, RemainDispCol - 5) >= 0, Cells(rowCurr, RemainDispCol - 2) - Cells(rowCurr, RemainDispCol - 5), 0)
'
'        Cells(rowCurr, ElasticCol) = Cells(rowCurr, ElasticCol - 4) - Cells(rowCurr, ElasticCol - 1)
'        Cells(rowCurr, CheckoutCoffCol) = Cells(rowCurr, CheckoutCoffCol - 2) / Cells(rowCurr, CheckoutCoffCol - 1)
'        Cells(rowCurr, RefRemainDispCol) = Cells(rowCurr, RefRemainDisp - 4) / Cells(rowCurr, RefRemainDispCol - 7)
'
'        rowCurr = rowCurr + 1
'    Wend

 
End Sub
'''������С/���У��ϵ���������Բ�����εȲ���
'Sub CalcParas()
'    Dim nWCs As Integer    '������
'    Dim nPs(10) As Integer    '�������������
'
'    nPN = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ")
'    Dim i, j As Integer
'    nWCs = Cells(1, 2)
'    For i = 0 To nWCs - 1
'        nPs(i) = Cells(2, 2 * (i + 1))
'    Next
'
'    Dim rowCurr As Integer    '��ָ��
'
'
'
'    For i = 0 To nWCs - 1    '��������
'        rowCurr = 3    'ÿ��������ʼ����ָ�����¿�ʼ
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
'''���ɸ���������Ӧ��
Private Sub GenerateRows_Click()
    Dim nWCs As Integer    '������
    Dim nPs(10) As Integer    '�������������
    Dim nPN     '����������Ӧ��������
    nPN = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ")
    Dim i, j As Integer
    nWCs = Cells(1, 2)
    For i = 0 To nWCs - 1
        nPs(i) = Cells(2, 2 * (i + 1))
        'Debug.Print nPs(i)
    Next
    'Debug.Print nWCs
    
    
    Dim rowCurr As Integer    '��ָ��
    rowCurr = First_Row
    
    For i = 0 To nWCs - 1    '��������
        For j = 1 To nPs(i)    '�������������Ĳ��
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
    Dim dispSummary(1 To 10) As String    '��word����ӦDocVariable��Ӧ
    Dim dispSummaryVar    '��word����ӦDocVariable��Ӧ
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

    For i = 1 To nWCs    'i��λ����
    
        '�������
        dispSummary(i) = "����" & CStr(nPN(i - 1)) & "���Խ������Ӷȼ���������x-x��ͼx-x������������������������Ӷ�У��ϵ����" & Format(StatPara(i, 1), "Fixed") & "��" & Format(StatPara(i, 2), "Fixed") & "֮�䣬" _
        & "���㡶��·���������������������̡��й涨��У��ϵ��С��1.0��Ҫ�����������������Բ������Ϊ" & Format(StatPara(i, 3), "Percent") & "��" _
        & "���㡶��·���������������������̡��й涨�Ĳ��������ֵҪ��(��ֵ20%)���ָ�״�����á�"
        
        wordApp.ActiveDocument.Variables(dispSummaryVar(i - 1)).value = dispSummary(i)
        wordApp.ActiveDocument.Fields.Update
        
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
        
        With tbl
            With .Borders
                .InsideLineStyle = wdLineStyleSingle
                .OutsideLineStyle = wdLineStyleSingle
            End With
        End With
    Next i
    
    'Dim TotalDisp(1 To MAX_NWC, 1 To 100)   'TotalDisp(i,j)��ʾ��i����������j������ܱ���
    
    'Dim NodeName(1 To MAX_NWC, 1 To 100) As String  '���������������
    'Dim Delta(1 To MAX_NWC, 1 To 100)
    'Dim RemainDisp(1 To MAX_NWC, 1 To 100)
    'Dim ElasticDisp(1 To MAX_NWC, 1 To 100)
    'Dim CheckoutCoff(1 To MAX_NWC, 1 To 100)
    'Dim RefRemainDisp(1 To MAX_NWC, 1 To 100)
    'Dim DispUbound(1 To MAX_NWC) As Integer    'ÿ�������Ͻ磨�½�Ϊ1��
    '
    'Dim StatPara(1 To MAX_NWC, 1 To 3)  'ͳ�Ʋ���,��СУ��ϵ�������У��ϵ���������Բ���Ӧ��
    ''StatPara(i,1~3)�ֱ��ʾ��i��������СУ��ϵ�������У��ϵ���������Բ���Ӧ��
    '
    'Dim t
    

    


    '���Կ��У�����ǩ���������
    'wordApp.ActiveDocument.Bookmarks("dispSummary1").Range.InsertAfter "111"
    

    '���Բ��������
    'wordApp.ActiveDocument.Tables.Add wordApp.ActiveDocument.Bookmarks("dispSummary1").Range, NumRows:=14 + 1, NumColumns:=7
    
    'Debug.Print wordApp.ActiveDocument.Variables("tb1").value
    
    
    'wordApp.Selection.Paste    '�������������ճ����"Select"�Ĳ���
    
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
    Dim wCondition(100)   '����wCondition(0)='һ',wCondition(1)='��'
    Dim wPointer As Integer
    
    Dim i As Integer
    Dim rowCurr As Integer    '��ָ��
    
    For i = 0 To 99
        wCondition(i) = ""
    Next
    
    First_Row = 3
    rowCurr = First_Row
    
    i = 0
    '�жϹ����������ظ��������㲻ͬ�Ĺ�����
    While Cells(rowCurr, 1) <> ""
        If i = 0 Then
            wCondition(i) = Cells(rowCurr, 1)
        ElseIf Cells(rowCurr, 1) <> wCondition(i) Then    '���¹���
            i = i + 1
            wCondition(i) = Cells(rowCurr, 1)
        End If
        
        rowCurr = rowCurr + 1
    Wend
    
    Dim nWK    '������
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
    tbl.Cell(1, 1).Range.InsertAfter "����"
    tbl.Cell(1, 2).Range.InsertAfter "�ܱ���"
    tbl.Cell(1, 3).Range.InsertAfter "���Ա���"
    tbl.Cell(1, 4).Range.InsertAfter "�������"
    tbl.Cell(1, 5).Range.InsertAfter "��������ֵ(mm)"
    tbl.Cell(1, 6).Range.InsertAfter "У��ϵ��"
    tbl.Cell(1, 7).Range.InsertAfter "��Բ������(%)"
    
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
    
    'Set t = wordApp.ActiveDocument.Tables(1)    '�����Ŵ�1��ʼ�㣿
    'tbl.Cell(1, 1).Range.InsertAfter "��һ����Ԫ��" '��һ����Ԫ����д���ַ�"��һ����Ԫ��"
    'tbl.Cell(tbl.Rows.Count, tbl.Columns.Count).Range.InsertAfter "���һ����Ԫ��" '�ڶ�����Ԫ����д���ַ�"���һ����Ԫ��"

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


