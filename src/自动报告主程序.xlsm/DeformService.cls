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

'nWCs:工况数
'nPs():每个工况测点数
'rowCurr:起始行
'WC_Col:工况列
'NodeName_Col:测点名列
'colorArray:单元格底色设置定位数组
Public Sub GenerateRows(ByVal NWCs As Integer, ByRef nPs() As Integer, ByVal rowCurr As Integer, ByVal WC_Col As Integer, ByVal NodeName_Col As Integer, Optional colorArray)

    Dim nPN     '各个工况对应中文名称
    nPN = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十")
    Dim i, j, k As Integer

'    For i = 0 To nWCs - 1
'        nPs(i) = Cells(2, 2 * (i + 1))
'    Next

    
    For i = 0 To NWCs - 1    '遍历工况
        For j = 1 To nPs(i)    '遍历各个工况的测点
            Cells(rowCurr, WC_Col) = nPN(i)
            Cells(rowCurr, NodeName_Col).NumberFormatLocal = "0.00"   '修改测点名列的格式，防止作图出现错误
            If Not IsMissing(colorArray) Then    '如果有该数组，则改变颜色
                For k = 1 To UBound(colorArray)
                    Cells(rowCurr, colorArray(k)).Interior.Color = RGB(0, 176, 80)
                Next
            End If
            rowCurr = rowCurr + 1
        Next
    Next
End Sub

'计算应变
'r2:变化后模数，r1:变化前模数，t2:变化后温度，t1：变化前温度
'instrumentType：仪器类型
'1:基康
'2:长沙金码
Public Function GetStrain(ByVal r2, ByVal r1, ByVal t2, ByVal t1, Optional instrumentType As Integer = 1) As Double
    Dim G, k, c
    G = 3.7: k = 1.8: c = 1.020019
    If instrumentType = 1 Then
        GetStrain = G * c * (r2 - r1) + k * (t2 - t1)
    ElseIf instrumentType = 2 Then
        GetStrain = r2 - r1
    Else    '异常则按基康计算
        GetStrain = G * c * (r2 - r1) + k * (t2 - t1)
    End If
End Function

'计算残余位移
'd0:初始位移;d1:满载位移;d2:卸载位移
'TODO:
'测试：
'1、0.00,3.00,0.3=>0.1
'2、0.00,3.00,-0.3=>0
'3、0.00,-3.00,-0.3=>0.1
'4、0.00,-3.00,0.3=>0
Public Function GetRemainDisp(ByVal d0, ByVal d1, ByVal d2) As Double
    If d1 - d0 >= 0 Then    '处理下挠工况
        GetRemainDisp = IIf(d2 - d0 >= 0, d2 - d0, 0)    '下挠工况卸载后若反弹至上挠，取0
    Else    'd1-d0<0: 处理上挠工况
        GetRemainDisp = IIf(d2 - d0 < 0, d2 - d0, 0)    '上挠工况卸载后若反弹至下挠，取0
    End If
End Function


