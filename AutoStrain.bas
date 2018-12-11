Attribute VB_Name = "AutoStrain"
Option Explicit
Dim FirstRow As Integer    '起始数据行数
Dim c1 As Integer    '满载模数计算（所在列，以下略）
Dim c2 As Integer    '满载温度计算
Dim c3 As Integer    '满载应变
Dim c4 As Integer    '卸载模数计算
Dim c5 As Integer    '卸载温度计算
Dim c6 As Integer    '卸载残余应变
Dim c7 As Integer    '理论应变

Dim d1 As Integer    '实测总应变（所在列，以下略）
Dim d2 As Integer    '弹性应变
Dim d3 As Integer    '残余应变
Dim d4 As Integer    '满载理论值
Dim d5 As Integer    '校验系数
Dim d6 As Integer    '相对残余应变
Sub AutoStrain_Click()

    Dim rowCurr As Integer    '行指针
    FirstRow = 10
    
    c1 = 14: c2 = 15: c3 = 16: c4 = 17: c5 = 18: c6 = 19: c7 = 20
    
    
    rowCurr = FirstRow
    
    While Cells(rowCurr, 1) <> ""
        Cells(rowCurr, TotalDispCol) = Cells(rowCurr, TotalDispCol - 1) - Cells(rowCurr, TotalDispCol - 2)
        rowCurr = rowCurr + 1
    Wend
   
End Sub
