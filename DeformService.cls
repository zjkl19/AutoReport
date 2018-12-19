VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeformService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'nWCs:������
'nPs():ÿ�����������
'rowCurr:��ʼ��
'WC_Col:������
Public Sub GenerateRows(ByVal nWCs As Integer, ByRef nPs() As Integer, ByVal rowCurr As Integer, ByVal WC_Col As Integer)

    Dim nPN     '����������Ӧ��������
    nPN = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ")
    Dim i, j As Integer

'    For i = 0 To nWCs - 1
'        nPs(i) = Cells(2, 2 * (i + 1))
'    Next

    
    For i = 0 To nWCs - 1    '��������
        For j = 1 To nPs(i)    '�������������Ĳ��
            Cells(rowCurr, WC_Col) = nPN(i)
            rowCurr = rowCurr + 1
        Next
    Next
End Sub

'����Ӧ��
'r2:�仯��ģ����r1:�仯ǰģ����t2:�仯���¶ȣ�t1���仯ǰ�¶�
Public Function GetStrain(ByVal r2, ByVal r1, ByVal t2, ByVal t1) As Double
    Dim G, K, C
    G = 3.7: K = 1.8: C = 1.020019
    GetStrain = G * C * (r2 - r1) + K * (t2 - t1)
End Function
