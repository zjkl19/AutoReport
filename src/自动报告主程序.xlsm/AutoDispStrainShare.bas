Attribute VB_Name = "AutoDispStrainShare"
Public Const DispSheetName As String = "�Ӷ�"
Public Const StrainSheetName As String = "Ӧ��"
 
Public DispTheoryShapeObjArray(1 To MAX_NWC) As Shape    '�洢������ֵͼָ��
Public DispTheoryShapeObjCounts As Integer


Public StrainTheoryShapeObjArray(1 To MAX_NWC) As Shape    '�洢������ֵͼָ��
Public StrainTheoryShapeObjCounts As Integer

Public nPN    '����������Ӧ��������

Public Const ExportPromteString As String = "������ɣ�����У���Զ������Ľ������ֹ����"
Public Const FontName_KaiGB2312 As String = "����_GB2312"


'ɾ����ǰsheet������ChartObject
Sub DeleteAllCharts()
    Dim c As ChartObject
    For Each c In ActiveSheet.ChartObjects
        c.Delete
    Next
End Sub
