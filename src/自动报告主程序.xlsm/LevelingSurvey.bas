Attribute VB_Name = "LevelingSurvey"
Option Explicit
Private Const MeasurePointName_Col As Integer = 1   '�����������
Private Const BacksightPointName_Col As Integer = 2
Private Const ForesightPoint_Col As Integer = 3
Private Const BacksightPoint_Col As Integer = 4
Private Const Altitude_Col As Integer = 5
Private Const TransData_Col As Integer = 6    '���㣨�ٷֱ�����������

Private Const First_Row As Integer = 2     '��ʼ��������

'�������
Public Sub LevelingSurveyDataClear()
  If (MsgBox("����������ݲ��ɳ�������ȷ��Ҫ�����", vbYesNo + vbExclamation, "�ò������ɳ���") = vbNo) Then
    Exit Sub
  End If
  
  Dim i, j, k As Integer
  Dim rowCurr As Integer    '��ָ��
  rowCurr = First_Row
  
  'TODO:�����ʼ�������
  Dim dataArray(1 To 100) As Integer    '����յ���
  k = 1
  For i = 1 To 6
    dataArray(k) = i
    k = k + 1
  Next i


  
  '��ձ������
  While Cells(rowCurr, 1) <> ""    '��һ����Ԫ��������Ϊ�ж�����
    For i = 1 To UBound(dataArray)
        If dataArray(i) = 0 Then Exit For
        Cells(rowCurr, dataArray(i)) = ""
    Next
    rowCurr = rowCurr + 1
  Wend

End Sub

'�Զ�ˮ׼����ת��
Private Sub AutoLevelingSurveyTransform()
    Sheets("ˮ׼����ת��").Activate
    
    Dim rowCurr As Integer
    
    Dim nP As Integer    '��������Զ����㣩

    Dim levelingData 'ˮ׼�������ݣ������,�߳�
    Set levelingData = CreateObject("Scripting.Dictionary")
    
    rowCurr = 2
    While Cells(rowCurr, MeasurePointName_Col) <> ""    '�Ƿ���ڼ�¼��
        If Cells(rowCurr, Altitude_Col) <> "" And rowCurr = 2 Then  '�߳��Ƿ���֪�����ǵ�2�У������и̶߳�Ҫ���ݵ�2�и߳������㣬�������ì�ܣ�
            levelingData.Add CStr(Cells(rowCurr, MeasurePointName_Col)), CDbl(Cells(rowCurr, Altitude_Col))
        Else    '��Ϊ��֪�򣺸߳�=����߳�+����-ǰ��
            levelingData.Add CStr(Cells(rowCurr, MeasurePointName_Col)), CDbl(levelingData.Item(CStr(Cells(rowCurr, BacksightPointName_Col)))) + CDbl(Cells(rowCurr, BacksightPoint_Col)) - CDbl(Cells(rowCurr, ForesightPoint_Col))
        End If
        
        If Cells(rowCurr, Altitude_Col) = "" Then    'д��δ֪�߳�����
            Cells(rowCurr, Altitude_Col) = CDbl(levelingData.Item(CStr(Cells(rowCurr, MeasurePointName_Col))))
        End If
        
        Cells(rowCurr, TransData_Col) = -1 * Cells(rowCurr, Altitude_Col) * 1000  '��λ��mm
        rowCurr = rowCurr + 1
    Wend
   

End Sub
