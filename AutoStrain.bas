Attribute VB_Name = "AutoStrain"
Option Explicit
Dim FirstRow As Integer    '��ʼ��������
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
Sub AutoStrain_Click()

    Dim rowCurr As Integer    '��ָ��
    FirstRow = 10
    
    c1 = 14: c2 = 15: c3 = 16: c4 = 17: c5 = 18: c6 = 19: c7 = 20
    
    
    rowCurr = FirstRow
    
    While Cells(rowCurr, 1) <> ""
        Cells(rowCurr, TotalDispCol) = Cells(rowCurr, TotalDispCol - 1) - Cells(rowCurr, TotalDispCol - 2)
        rowCurr = rowCurr + 1
    Wend
   
End Sub
