Attribute VB_Name = "AutoThrust"
Option Explicit

Public nThrust As Integer    '栏杆推力测点数
Public ThrustLevel As Integer     '栏杆推力加载级数

Public Const MAX_nThrust As Integer = 100  '栏杆推力最多点数
Public Const MAX_ThrustLevel As Integer = 10     '栏杆推力最大加载级数

Private Const First_Row As Integer = 13     '起始数据行数


Private Const ThrustNodeName_Col As Integer = 1
Private Const ThrustLevel_Col As Integer = 2
Private Const ThrustTotalDisp_Col As Integer = 3
Private Const ThrustElasticDisp_Col As Integer = 4
Private Const ThrustRemainDisp_Col As Integer = 5
Private Const ThrustRefRemainDisp_Col As Integer = 6

Public ThrustNodeName(1 To MAX_nThrust, 1 To 100) As String  '各个工况测点名称
Public ThrustTotalDisp(1 To MAX_nThrust, 1 To MAX_ThrustLevel)   'ThrustTotalDisp(i,j)表示第i个测点，第j级总变形（最后一级表示退载）
Public ThrustElasticDisp(1 To MAX_nThrust)    '各个测点弹性变形
Public ThrustRemainDisp(1 To MAX_nThrust)    '各个测点残余变形
Public ThrustRefRemainDisp(1 To MAX_nThrust)    '各个测点相对残余变形


'变量初始化
Private Sub InitVar()
    nThrust = Cells(1, 2)
    ThrustLevel = Cells(2, 2)
End Sub

'栏杆推力自动计算
Public Sub AutoThrust()

    InitVar
    
    Dim rowCurr As Integer    '行指针
    
    Dim i, j As Integer
    rowCurr = First_Row
    For i = 1 To nThrust
        For j = 1 To ThrustLevel + 1
        
            ThrustTotalDisp(i, j) = Cells(rowCurr, ThrustTotalDisp_Col)
            ThrustTotalDisp(i, j) = Round(ThrustTotalDisp(i, j), 2)
            
            Cells(rowCurr, ThrustElasticDisp_Col) = "/"    '加载未满级的弹性变形
            Range(Cells(rowCurr, ThrustElasticDisp_Col), Cells(rowCurr, ThrustElasticDisp_Col)).HorizontalAlignment = xlCenter
            
            If j = ThrustLevel + 1 Then    '获得退载数据后，可计算弹性变形
                ThrustElasticDisp(i) = ThrustTotalDisp(i, j - 1) - ThrustTotalDisp(i, j)  '上一级变形-残余变形（退载值）
                

                Cells(rowCurr - 1, ThrustElasticDisp_Col) = ThrustElasticDisp(i)
                Cells(rowCurr, ThrustElasticDisp_Col) = "/"
                
                ThrustRemainDisp(i) = ThrustTotalDisp(i, j)
                Cells(rowCurr - ThrustLevel, ThrustRemainDisp_Col) = ThrustRemainDisp(i)
                
                ThrustRefRemainDisp(i) = ThrustRemainDisp(i) / ThrustTotalDisp(i, j - 1)
                Cells(rowCurr - ThrustLevel, ThrustRefRemainDisp_Col) = Format(ThrustRefRemainDisp(i), "Percent")
                
                Range(Cells(rowCurr, ThrustRemainDisp_Col), Cells(rowCurr - ThrustLevel, ThrustRemainDisp_Col)).Merge   '合并残余变形
                Range(Cells(rowCurr, ThrustRemainDisp_Col), Cells(rowCurr - ThrustLevel, ThrustRemainDisp_Col)).HorizontalAlignment = xlCenter ''左右居中
                Range(Cells(rowCurr, ThrustRemainDisp_Col), Cells(rowCurr - ThrustLevel, ThrustRemainDisp_Col)).VerticalAlignment = xlCenter ''上下居中
                
                Range(Cells(rowCurr, ThrustRefRemainDisp_Col), Cells(rowCurr - ThrustLevel, ThrustRefRemainDisp_Col)).Merge   '合并相对残余变形
                Range(Cells(rowCurr, ThrustRefRemainDisp_Col), Cells(rowCurr - ThrustLevel, ThrustRefRemainDisp_Col)).HorizontalAlignment = xlCenter ''左右居中
                Range(Cells(rowCurr, ThrustRefRemainDisp_Col), Cells(rowCurr - ThrustLevel, ThrustRefRemainDisp_Col)).VerticalAlignment = xlCenter ''上下居中
            End If
            
            rowCurr = rowCurr + 1
        Next
    Next
 
End Sub

'生成栏杆推力行
Public Sub GenerateThrustRows()
    InitVar    '获取测点数
    
    Dim i, j As Integer
    Dim rowCurr As Integer    '行指针
    Dim bgColor
    bgColor = RGB(0, 176, 80)
    rowCurr = First_Row
    
    
    For i = 1 To nThrust    '遍历测点
        For j = 1 To ThrustLevel + 1  '遍历各级
            
            Cells(rowCurr, ThrustNodeName_Col) = CStr(i) & "#"
            If j <> ThrustLevel + 1 Then
                Cells(rowCurr, ThrustLevel_Col) = CStr(j) & "级"
            Else
                Cells(rowCurr, ThrustLevel_Col) = "退载"
            End If
            
            '设置必填项的背景色
            Cells(rowCurr, ThrustNodeName_Col).Interior.Color = bgColor
            Cells(rowCurr, ThrustLevel_Col).Interior.Color = bgColor
            Cells(rowCurr, ThrustTotalDisp_Col).Interior.Color = bgColor
            
            rowCurr = rowCurr + 1
        Next
    Next

End Sub

'清空数据
Public Sub ThrustDataClear()
  If (MsgBox("清空输入数据不可撤销，你确定要清空吗？", vbYesNo + vbExclamation, "该操作不可撤销") = vbNo) Then
    Exit Sub
  End If
  
  Dim i, j As Integer
  Dim rowCurr As Integer    '行指针
  rowCurr = First_Row
  
  '清空表格数据
  While Cells(rowCurr, 1) <> ""    '第一个单元格数据作为判断依据
    For i = 1 To ThrustRefRemainDisp_Col
        Cells(rowCurr, i) = ""
        Cells(rowCurr, i).Interior.Color = RGB(255, 255, 255) ' RGB(0, 176, 80)
    Next
    rowCurr = rowCurr + 1
  Wend

End Sub
