Attribute VB_Name = "AutoThrust"
Option Explicit

Public nThrust As Integer    '�������������
Public ThrustLevel As Integer     '�����������ؼ���

Public Const MAX_nThrust As Integer = 100  '��������������
Public Const MAX_ThrustLevel As Integer = 10     '�������������ؼ���

Private Const First_Row As Integer = 13     '��ʼ��������


Private Const ThrustNodeName_Col As Integer = 1
Private Const ThrustLevel_Col As Integer = 2
Private Const ThrustTotalDisp_Col As Integer = 3
Private Const ThrustElasticDisp_Col As Integer = 4
Private Const ThrustRemainDisp_Col As Integer = 5
Private Const ThrustRefRemainDisp_Col As Integer = 6

Public ThrustNodeName(1 To MAX_nThrust, 1 To 100) As String  '���������������
Public ThrustTotalDisp(1 To MAX_nThrust, 1 To MAX_ThrustLevel)   'ThrustTotalDisp(i,j)��ʾ��i����㣬��j���ܱ��Σ����һ����ʾ���أ�
Public ThrustElasticDisp(1 To MAX_nThrust)    '������㵯�Ա���
Public ThrustRemainDisp(1 To MAX_nThrust)    '�������������
Public ThrustRefRemainDisp(1 To MAX_nThrust)    '���������Բ������


'������ʼ��
Private Sub InitVar()
    nThrust = Cells(1, 2)
    ThrustLevel = Cells(2, 2)
End Sub

'���������Զ�����
Public Sub AutoThrust()

    InitVar
    
    Dim rowCurr As Integer    '��ָ��
    
    Dim i, j As Integer
    rowCurr = First_Row
    For i = 1 To nThrust
        For j = 1 To ThrustLevel + 1
        
            ThrustTotalDisp(i, j) = Cells(rowCurr, ThrustTotalDisp_Col)
            ThrustTotalDisp(i, j) = Round(ThrustTotalDisp(i, j), 2)
            
            Cells(rowCurr, ThrustElasticDisp_Col) = "/"    '����δ�����ĵ��Ա���
            Range(Cells(rowCurr, ThrustElasticDisp_Col), Cells(rowCurr, ThrustElasticDisp_Col)).HorizontalAlignment = xlCenter
            
            If j = ThrustLevel + 1 Then    '����������ݺ󣬿ɼ��㵯�Ա���
                ThrustElasticDisp(i) = ThrustTotalDisp(i, j - 1) - ThrustTotalDisp(i, j)  '��һ������-������Σ�����ֵ��
                

                Cells(rowCurr - 1, ThrustElasticDisp_Col) = ThrustElasticDisp(i)
                Cells(rowCurr, ThrustElasticDisp_Col) = "/"
                
                ThrustRemainDisp(i) = ThrustTotalDisp(i, j)
                Cells(rowCurr - ThrustLevel, ThrustRemainDisp_Col) = ThrustRemainDisp(i)
                
                ThrustRefRemainDisp(i) = ThrustRemainDisp(i) / ThrustTotalDisp(i, j - 1)
                Cells(rowCurr - ThrustLevel, ThrustRefRemainDisp_Col) = Format(ThrustRefRemainDisp(i), "Percent")
                
                Range(Cells(rowCurr, ThrustRemainDisp_Col), Cells(rowCurr - ThrustLevel, ThrustRemainDisp_Col)).Merge   '�ϲ��������
                Range(Cells(rowCurr, ThrustRemainDisp_Col), Cells(rowCurr - ThrustLevel, ThrustRemainDisp_Col)).HorizontalAlignment = xlCenter ''���Ҿ���
                Range(Cells(rowCurr, ThrustRemainDisp_Col), Cells(rowCurr - ThrustLevel, ThrustRemainDisp_Col)).VerticalAlignment = xlCenter ''���¾���
                
                Range(Cells(rowCurr, ThrustRefRemainDisp_Col), Cells(rowCurr - ThrustLevel, ThrustRefRemainDisp_Col)).Merge   '�ϲ���Բ������
                Range(Cells(rowCurr, ThrustRefRemainDisp_Col), Cells(rowCurr - ThrustLevel, ThrustRefRemainDisp_Col)).HorizontalAlignment = xlCenter ''���Ҿ���
                Range(Cells(rowCurr, ThrustRefRemainDisp_Col), Cells(rowCurr - ThrustLevel, ThrustRefRemainDisp_Col)).VerticalAlignment = xlCenter ''���¾���
            End If
            
            rowCurr = rowCurr + 1
        Next
    Next
 
End Sub

'��������������
Public Sub GenerateThrustRows()
    InitVar    '��ȡ�����
    
    Dim i, j As Integer
    Dim rowCurr As Integer    '��ָ��
    Dim bgColor
    bgColor = RGB(0, 176, 80)
    rowCurr = First_Row
    
    
    For i = 1 To nThrust    '�������
        For j = 1 To ThrustLevel + 1  '��������
            
            Cells(rowCurr, ThrustNodeName_Col) = CStr(i) & "#"
            If j <> ThrustLevel + 1 Then
                Cells(rowCurr, ThrustLevel_Col) = CStr(j) & "��"
            Else
                Cells(rowCurr, ThrustLevel_Col) = "����"
            End If
            
            '���ñ�����ı���ɫ
            Cells(rowCurr, ThrustNodeName_Col).Interior.Color = bgColor
            Cells(rowCurr, ThrustLevel_Col).Interior.Color = bgColor
            Cells(rowCurr, ThrustTotalDisp_Col).Interior.Color = bgColor
            
            rowCurr = rowCurr + 1
        Next
    Next

End Sub

'�������
Public Sub ThrustDataClear()
  If (MsgBox("����������ݲ��ɳ�������ȷ��Ҫ�����", vbYesNo + vbExclamation, "�ò������ɳ���") = vbNo) Then
    Exit Sub
  End If
  
  Dim i, j As Integer
  Dim rowCurr As Integer    '��ָ��
  rowCurr = First_Row
  
  '��ձ������
  While Cells(rowCurr, 1) <> ""    '��һ����Ԫ��������Ϊ�ж�����
    For i = 1 To ThrustRefRemainDisp_Col
        Cells(rowCurr, i) = ""
        Cells(rowCurr, i).Interior.Color = RGB(255, 255, 255) ' RGB(0, 176, 80)
    Next
    rowCurr = rowCurr + 1
  Wend

End Sub
