Attribute VB_Name = "AutoStrain"
Option Explicit

Private Const e1 As Integer = 3  '��ʼģ���������У������ԣ�
Private Const e2 As Integer = 4  '��ʼ�¶�
Private Const e3 As Integer = 5  '����ģ��
Private Const e4 As Integer = 6  '�����¶�
Private Const e5 As Integer = 7  'ж��ģ��
Private Const e6 As Integer = 8  'ж���¶�

Dim c1 As Integer    '����ģ�����㣨�����У������ԣ�
Dim c2 As Integer    '�����¶ȼ���
Dim c3 As Integer    '����Ӧ��
Dim c4 As Integer    'ж��ģ������
Dim c5 As Integer    'ж���¶ȼ���
Dim c6 As Integer    'ж�ز���Ӧ��
Dim c7 As Integer    '����Ӧ��

Dim d1 As Integer    'ʵ����Ӧ�䣨�����У������ԣ�
Dim d2 As Integer    '����Ӧ��
Dim d3 As Integer    '����Ӧ��
Dim d4 As Integer    '��������ֵ
Dim d5 As Integer    'У��ϵ��
Dim d6 As Integer    '��Բ���Ӧ��

Private Const First_Row As Integer = 13    '��ʼ��������
Private Const StrainStatPara1_Row As Integer = 4
Private Const StrainStatPara2_Row As Integer = 5
Private Const StrainStatPara3_Row As Integer = 6

Const StrainNode_Name_Col As Integer = 2  '�����������

Public StrainGlobalWC(1 To MAX_NWC)    'ȫ�ֹ�����λ����

Public StrainNodeName(1 To MAX_NWC, 1 To MAX_NPS) As String  '���������������
Public TotalStrain(1 To MAX_NWC, 1 To MAX_NPS)    '����Ӧ��
Public RemainStrain(1 To MAX_NWC, 1 To MAX_NPS)    '����Ӧ�䣨����Ӧ�䣩
Public ElasticStrain(1 To MAX_NWC, 1 To MAX_NPS)
Public TheoryStrain(1 To MAX_NWC, 1 To MAX_NPS)
Public StrainCheckoutCoff(1 To MAX_NWC, 1 To MAX_NPS)
Public RefRemainStrain(1 To MAX_NWC, 1 To MAX_NPS)

Public StrainStatPara(1 To MAX_NWC, 1 To 3)  'ͳ�Ʋ���,��СУ��ϵ�������У��ϵ���������Բ���Ӧ��

Private Const TotalStrain_Col As Integer = 27
Private Const RemainStrain_Col As Integer = 29
Private Const ElasticStrain_Col As Integer = 28
Private Const TheoryStrain_Col As Integer = 30
Private Const StrainCheckoutCoff_Col As Integer = 31
Private Const RefRemainStrain_Col As Integer = 32

Public StrainUbound(1 To MAX_NWC) As Integer    'ÿ�������Ͻ磨�½�Ϊ1��
Public StrainNWCs As Integer    'Ӧ�乤����
Public StrainNPs(10) As Integer    '�������������
Public Sub InitStrainVar()
    Dim i As Integer
    
    StrainNWCs = Cells(1, 2)
    For i = 1 To StrainNWCs
        StrainUbound(i) = Cells(2, 2 * i)
    Next
 
    For i = 1 To StrainNWCs
        StrainGlobalWC(i) = Cells(3, 2 * i)
    Next
End Sub
'����Ӧ�������
Private Sub GenerateStrainRows_Click()

    StrainNWCs = Cells(1, 2)
    
    Dim i As Integer
    For i = 0 To StrainNWCs - 1    'ÿ�����������
        StrainNPs(i) = Cells(2, 2 * (i + 1))
    Next
   
    Dim rowCurr As Integer    '��ָ��
    rowCurr = First_Row
    
    Dim ds As DeformService
    Set ds = New DeformService
    ds.GenerateRows StrainNWCs, StrainNPs, rowCurr, 1
    Set ds = Nothing
 
End Sub

'����Ӧ��
'r2:�仯��ģ����r1:�仯ǰģ����t2:�仯���¶ȣ�t1���仯ǰ�¶�
Public Function GetStrain(ByVal r2, ByVal r1, ByVal t2, ByVal t1)
    Dim G, K, C
    G = 3.7: K = 1.8: C = 1.020019
    GetStrain = G * C * (r2 - r1) + K * (t2 - t1)
End Function

'�Զ�����Ӧ��
Public Sub AutoStrain_Click()

    InitStrainVar
    
    Dim i, j As Integer

    Dim rowCurr As Integer    '��ָ��
    rowCurr = First_Row
    
    For i = 1 To StrainNWCs
        For j = 1 To StrainUbound(i)
        
            StrainNodeName(i, j) = Cells(rowCurr, StrainNode_Name_Col)
            
            TotalStrain(i, j) = GetStrain(Cells(rowCurr, e3), Cells(rowCurr, e1), Cells(rowCurr, e4), Cells(rowCurr, e2))
            Cells(rowCurr, TotalStrain_Col) = TotalStrain(i, j)
            
             '�㷨��ж�����ʼ��ֵ>=0��ȡж�����ʼ��ֵ������ȡ0
            RemainStrain(i, j) = IIf(GetStrain(Cells(rowCurr, e5), Cells(rowCurr, e1), Cells(rowCurr, e6), Cells(rowCurr, e2)) >= 0 _
            , GetStrain(Cells(rowCurr, e5), Cells(rowCurr, e1), Cells(rowCurr, e6), Cells(rowCurr, e2)), 0)
            Cells(rowCurr, RemainStrain_Col) = RemainStrain(i, j)    '����Ӧ��
            
            ElasticStrain(i, j) = TotalStrain(i, j) - RemainStrain(i, j)
            Cells(rowCurr, ElasticStrain_Col) = ElasticStrain(i, j)    '����Ӧ��
             
            TheoryStrain(i, j) = Cells(rowCurr, TheoryStrain_Col)    '����Ӧ��ֱ��ȡֵ
            
            StrainCheckoutCoff(i, j) = ElasticStrain(i, j) / TheoryStrain(i, j)
            Cells(rowCurr, StrainCheckoutCoff_Col) = StrainCheckoutCoff(i, j)    'У��ϵ��
             
            RefRemainStrain(i, j) = RemainStrain(i, j) / TotalStrain(i, j)
            Cells(rowCurr, RefRemainStrain_Col) = RefRemainStrain(i, j)    '��Բ������
            
            rowCurr = rowCurr + 1
        Next
    Next
    
    '�������������С/��У��ϵ���������Բ���Ӧ��
    For i = 1 To StrainNWCs
        StrainStatPara(i, 1) = StrainCheckoutCoff(i, 1): StrainStatPara(i, 2) = StrainCheckoutCoff(i, 1): StrainStatPara(i, 3) = RefRemainStrain(i, 1)
        For j = 1 To StrainUbound(i)
            If (StrainCheckoutCoff(i, j) < StrainStatPara(i, 1)) Then
                StrainStatPara(i, 1) = StrainCheckoutCoff(i, j)
            End If
            If (StrainCheckoutCoff(i, j) > StrainStatPara(i, 2)) Then
                StrainStatPara(i, 2) = StrainCheckoutCoff(i, j)
            End If
            If (RefRemainStrain(i, j) > StrainStatPara(i, 3)) Then
                StrainStatPara(i, 3) = RefRemainStrain(i, j)
            End If
        Next

        '����д��Excel
        Cells(StrainStatPara1_Row, 2 * i) = Format(StrainStatPara(i, 1), "Fixed"): Cells(StrainStatPara2_Row, 2 * i) = Format(StrainStatPara(i, 2), "Fixed"): Cells(StrainStatPara3_Row, 2 * i) = Format(StrainStatPara(i, 3), "Percent")
    Next
 
End Sub

