Attribute VB_Name = "AutoDisp"
Option Explicit

Private Const First_Row As Integer = 11     '��ʼ��������
Private Const StatPara1_Row As Integer = 4
Private Const StatPara2_Row As Integer = 5
Private Const StatPara3_Row As Integer = 6

Const WC_Col As Integer = 1     '������������
Public Const MAX_NWC As Integer = 10     '��󹤿���
Public Const MAX_NPS As Integer = 100     'ÿ�������������

Const Node_Name_Col As Integer = 2  '�����������
Const TheoryDisp_Col As Integer = 10  '����λ��������
Dim TotalDispCol As Integer    '�ܱ���������
Dim DeltaCol As Integer   '����������
Dim RemainDispCol As Integer    '�������������
Dim ElasticCol As Integer    '���Ա���������
Dim CheckoutCoffCol As Integer    'У��ϵ��������
Dim RefRemainDispCol As Integer    '��Բ������������

Public nWCs As Integer    '���Ӷȣ�������
Public nPN    '����������Ӧ��������
'nPN = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ")

Public GlobalWC(1 To MAX_NWC)    'ȫ�ֹ�����λ����

Public TotalDisp(1 To MAX_NWC, 1 To 100)   'TotalDisp(i,j)��ʾ��i����������j������ܱ���
Public NodeName(1 To MAX_NWC, 1 To 100) As String  '���������������
Public Delta(1 To MAX_NWC, 1 To 100)    '���������ñ���
Public RemainDisp(1 To MAX_NWC, 1 To 100)
Public ElasticDisp(1 To MAX_NWC, 1 To 100)
Public TheoryDisp(1 To MAX_NWC, 1 To 100)
Public CheckoutCoff(1 To MAX_NWC, 1 To 100)
Public RefRemainDisp(1 To MAX_NWC, 1 To 100)
Public DispUbound(1 To MAX_NWC) As Integer    'ÿ�������Ͻ磨�½�Ϊ1��

Public StatPara(1 To MAX_NWC, 1 To 3)  'ͳ�Ʋ���,��СУ��ϵ�������У��ϵ���������Բ���Ӧ��
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
    

    For i = 1 To nWCs
        GlobalWC(i) = Cells(3, 2 * i)
    Next
    
    TotalDispCol = 5
    DeltaCol = 7
    RemainDispCol = 8
    ElasticCol = 9
    CheckoutCoffCol = 11
    RefRemainDispCol = 12
End Sub


Public Sub AutoDisp_Click()

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
        Cells(StatPara1_Row, 2 * i) = Format(StatPara(i, 1), "Fixed"): Cells(StatPara2_Row, 2 * i) = Format(StatPara(i, 2), "Fixed"): Cells(StatPara3_Row, 2 * i) = Format(StatPara(i, 3), "Percent")
    Next

 
End Sub

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


