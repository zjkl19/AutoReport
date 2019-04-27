Attribute VB_Name = "LevelingSurvey"
Option Explicit
Private Const MeasurePointName_Col As Integer = 1   '测点编号所在列
Private Const BacksightPointName_Col As Integer = 2
Private Const ForesightPoint_Col As Integer = 3
Private Const BacksightPoint_Col As Integer = 4
Private Const Altitude_Col As Integer = 5
Private Const TransData_Col As Integer = 6    '换算（百分表）读数所在列

Private Const First_Row As Integer = 2     '起始数据行数

'清空数据
Public Sub LevelingSurveyDataClear()
  If (MsgBox("清空输入数据不可撤销，你确定要清空吗？", vbYesNo + vbExclamation, "该操作不可撤销") = vbNo) Then
    Exit Sub
  End If
  
  Dim i, j, k As Integer
  Dim rowCurr As Integer    '行指针
  rowCurr = First_Row
  
  'TODO:数组初始化更灵活
  Dim dataArray(1 To 100) As Integer    '待清空的列
  k = 1
  For i = 1 To 6
    dataArray(k) = i
    k = k + 1
  Next i


  
  '清空表格数据
  While Cells(rowCurr, 1) <> ""    '第一个单元格数据作为判断依据
    For i = 1 To UBound(dataArray)
        If dataArray(i) = 0 Then Exit For
        Cells(rowCurr, dataArray(i)) = ""
    Next
    rowCurr = rowCurr + 1
  Wend

End Sub

'自动水准测量转换
Private Sub AutoLevelingSurveyTransform()
    Sheets("水准测量转换").Activate
    
    Dim rowCurr As Integer
    
    Dim nP As Integer    '测点数（自动计算）

    Dim levelingData '水准测量数据：测点编号,高程
    Set levelingData = CreateObject("Scripting.Dictionary")
    
    rowCurr = 2
    While Cells(rowCurr, MeasurePointName_Col) <> ""    '是否存在记录？
        If Cells(rowCurr, Altitude_Col) <> "" And rowCurr = 2 Then  '高程是否已知并且是第2行（其它行高程都要根据第2行高程来推算，否则会有矛盾）
            levelingData.Add CStr(Cells(rowCurr, MeasurePointName_Col)), CDbl(Cells(rowCurr, Altitude_Col))
        Else    '不为已知则：高程=基点高程+后视-前视
            levelingData.Add CStr(Cells(rowCurr, MeasurePointName_Col)), CDbl(levelingData.Item(CStr(Cells(rowCurr, BacksightPointName_Col)))) + CDbl(Cells(rowCurr, BacksightPoint_Col)) - CDbl(Cells(rowCurr, ForesightPoint_Col))
        End If
        
        If Cells(rowCurr, Altitude_Col) = "" Then    '写入未知高程数据
            Cells(rowCurr, Altitude_Col) = CDbl(levelingData.Item(CStr(Cells(rowCurr, MeasurePointName_Col))))
        End If
        
        Cells(rowCurr, TransData_Col) = -1 * Cells(rowCurr, Altitude_Col) * 1000  '单位：mm
        rowCurr = rowCurr + 1
    Wend
   

End Sub
