Attribute VB_Name = "AutoDisp"
Option Explicit
Const First_Row As Integer = 8     '��ʼ��������
Const WC_Col As Integer = 1     '������������
Const MAX_NWC As Integer = 10     '��󹤿���

Dim TotalDispCol As Integer    '�ܱ���������
Dim DeltaCol As Integer   '����������
Dim RemainDispCol As Integer    '�������������
Dim ElasticCol As Integer    '���Ա���������
Dim CheckoutCoffCol As Integer    'У��ϵ��������
Dim RefRemainDispCol As Integer    '��Բ������������

Dim nWCs As Integer    '������
Dim TotalDisp(1 To MAX_NWC, 1 To 100)   'TotalDisp(i,j)��ʾ��i����������j������ܱ���
Dim Delta
Dim RemainDisp
Dim ElasticDisp
Dim CheckoutCoff
Dim RefRemainDisp
Dim DispUbound(1 To MAX_NWC) As Integer

Dim t

'''��ʼ��ȫ�ֱ���
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
    
    Dim rowCurr As Integer    '��ָ��
    

    
    rowCurr = First_Row
    While Cells(rowCurr, 1) <> ""
        Cells(rowCurr, TotalDispCol) = Cells(rowCurr, TotalDispCol - 1) - Cells(rowCurr, TotalDispCol - 2)
        Cells(rowCurr, DeltaCol) = Cells(rowCurr, DeltaCol - 1) - Cells(rowCurr, DeltaCol - 3)
        
        '�㷨��ж�������ض�����ֵ>=0��ȡж�������ض�����ֵ������ȡ0
        Cells(rowCurr, RemainDispCol) = IIf(Cells(rowCurr, RemainDispCol - 2) - Cells(rowCurr, RemainDispCol - 5) >= 0, Cells(rowCurr, RemainDispCol - 2) - Cells(rowCurr, RemainDispCol - 5), 0)
        
        Cells(rowCurr, ElasticCol) = Cells(rowCurr, ElasticCol - 4) - Cells(rowCurr, ElasticCol - 1)
        Cells(rowCurr, CheckoutCoffCol) = Cells(rowCurr, CheckoutCoffCol - 2) / Cells(rowCurr, CheckoutCoffCol - 1)
        Cells(rowCurr, RefRemainDispCol) = Cells(rowCurr, RefRemainDisp - 4) / Cells(rowCurr, RefRemainDispCol - 7)
        
        rowCurr = rowCurr + 1
    Wend

 
End Sub
'''������С/���У��ϵ���������Բ�����εȲ���
Sub CalcParas()
    Dim nWCs As Integer    '������
    Dim nPs(10) As Integer    '�������������
    Dim nPN     '����������Ӧ��������
    nPN = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ")
    Dim i, j As Integer
    nWCs = Cells(1, 2)
    For i = 0 To nWCs - 1
        nPs(i) = Cells(2, 2 * (i + 1))
    Next
    
    Dim rowCurr As Integer    '��ָ��
    

    
    For i = 0 To nWCs - 1    '��������
        rowCurr = 3    'ÿ��������ʼ����ָ�����¿�ʼ
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
Sub test()
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


