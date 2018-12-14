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
Public Sub GenerateRows(ByVal nWCs As Integer, ByRef nPs() As Integer, ByVal rowCurr As Integer, ByVal WC_Col As Integer)

    Dim nPN     '各个工况对应中文名称
    nPN = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十")
    Dim i, j As Integer

'    For i = 0 To nWCs - 1
'        nPs(i) = Cells(2, 2 * (i + 1))
'    Next

    
    For i = 0 To nWCs - 1    '遍历工况
        For j = 1 To nPs(i)    '遍历各个工况的测点
            Cells(rowCurr, WC_Col) = nPN(i)
            rowCurr = rowCurr + 1
        Next
    Next
End Sub

'计算应变
'r2:变化后模数，r1:变化前模数，t2:变化后温度，t1：变化前温度
Public Function GetStrain(ByVal r2, ByVal r1, ByVal t2, ByVal t1) As Double
    Dim G, K, C
    G = 3.7: K = 1.8: C = 1.020019
    GetStrain = G * C * (r2 - r1) + K * (t2 - t1)
End Function

