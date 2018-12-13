Attribute VB_Name = "AutoStrain"
Option Explicit

Private Const e1 As Integer = 3  '初始模数（所在列，以下略）
Private Const e2 As Integer = 4  '初始温度
Private Const e3 As Integer = 5  '满载模数
Private Const e4 As Integer = 6  '满载温度
Private Const e5 As Integer = 7  '卸载模数
Private Const e6 As Integer = 8  '卸载温度



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

Private Const First_Row As Integer = 10     '起始数据行数
Const StrainNode_Name_Col As Integer = 2  '测点编号所在列

Dim StrainNodeName(1 To MAX_NWC, 1 To MAX_NPS) As String  '各个工况测点名称

Dim TotalStrain(1 To MAX_NWC, 1 To MAX_NPS)    '满载应变
Dim RemainStrain(1 To MAX_NWC, 1 To MAX_NPS)    '残余应变（残余应变）
Dim ElasticStrain(1 To MAX_NWC, 1 To MAX_NPS)
Dim TheoryStrain(1 To MAX_NWC, 1 To MAX_NPS)
Dim StrainCheckoutCoff(1 To MAX_NWC, 1 To MAX_NPS)
Dim RefRemainStrain(1 To MAX_NWC, 1 To MAX_NPS)

Dim StrainStatPara(1 To MAX_NWC, 1 To 3)  '统计参数,最小校验系数，最大校验系数，最大相对残余应变

Private Const TotalStrain_Col As Integer = 27
Private Const RemainStrain_Col As Integer = 29
Private Const ElasticStrain_Col As Integer = 28
Private Const TheoryStrain_Col As Integer = 30
Private Const StrainCheckoutCoff_Col As Integer = 31
Private Const RefRemainStrain_Col As Integer = 32

Dim StrainUbound(1 To MAX_NWC) As Integer    '每个工况上界（下界为1）
Dim nWCs As Integer    '工况数
Dim nPs(10) As Integer    '各个工况测点数
Public Sub InitStrainVar()
    Dim i As Integer
    
    nWCs = Cells(1, 2)
    For i = 1 To nWCs
        StrainUbound(i) = Cells(2, 2 * i)
    Next
 
End Sub
'生成应变表格的行
Private Sub GenerateStrainRows_Click()

    nWCs = Cells(1, 2)
    
    Dim i As Integer
    For i = 0 To nWCs - 1    '每个工况测点数
        nPs(i) = Cells(2, 2 * (i + 1))
    Next
   
    Dim rowCurr As Integer    '行指针
    rowCurr = First_Row
    
    Dim ds As DeformService
    Set ds = New DeformService
    ds.GenerateRows nWCs, nPs, rowCurr, 1
    Set ds = Nothing
 
End Sub

Public Function GetStrain(ByVal r2, ByVal r1, ByVal t2, ByVal t1)
    Dim G, K, C
    G = 3.7: K = 1.8: C = 1.020019
    GetStrain = G * C * (r2 - r1) + K * (t2 - t1)
End Function

'自动计算应变
Private Sub AutoStrain_Click()

    InitStrainVar
    
    Dim i, j As Integer

    Dim rowCurr As Integer    '行指针
    rowCurr = First_Row
    
    For i = 1 To nWCs
        For j = 1 To StrainUbound(i)
        
            StrainNodeName(i, j) = Cells(rowCurr, StrainNode_Name_Col)
            
            TotalStrain(i, j) = GetStrain(Cells(rowCurr, e3), Cells(rowCurr, e1), Cells(rowCurr, e4), Cells(rowCurr, e2))
            Cells(rowCurr, TotalStrain_Col) = TotalStrain(i, j)
            
             '算法：卸载与初始差值>=0，取卸载与初始差值，否则取0
            RemainStrain(i, j) = IIf(GetStrain(Cells(rowCurr, e5), Cells(rowCurr, e1), Cells(rowCurr, e6), Cells(rowCurr, e2)) >= 0 _
            , GetStrain(Cells(rowCurr, e5), Cells(rowCurr, e1), Cells(rowCurr, e6), Cells(rowCurr, e2)), 0)
            Cells(rowCurr, RemainStrain_Col) = RemainStrain(i, j)    '残余应变
            
            ElasticStrain(i, j) = TotalStrain(i, j) - RemainStrain(i, j)
            Cells(rowCurr, ElasticStrain_Col) = ElasticStrain(i, j)    '弹性应变
             
            TheoryStrain(i, j) = Cells(rowCurr, TheoryStrain_Col)    '理论应变直接取值
            
            StrainCheckoutCoff(i, j) = ElasticStrain(i, j) / TheoryStrain(i, j)
            Cells(rowCurr, StrainCheckoutCoff_Col) = StrainCheckoutCoff(i, j)    '校验系数
             
            RefRemainStrain(i, j) = RemainStrain(i, j) / TotalStrain(i, j)
            Cells(rowCurr, RefRemainStrain_Col) = RefRemainStrain(i, j)    '相对残余变形
            
            rowCurr = rowCurr + 1
        Next
    Next
    
    '计算各个工况最小/大校验系数，最大相对残余应变
    For i = 1 To nWCs
        StrainStatPara(i, 1) = StrainCheckoutCoff(i, 1): StrainStatPara(i, 2) = StrainCheckoutCoff(i, 1): StrainStatPara(i, 3) = RefRemainStrain(i, 1)
        For j = 1 To StrainUbound(i)
            If (StrainCheckoutCoff(i, j) < StrainStatPara(i, 1)) Then
                StrainStatPara(i, 1) = StrainCheckoutCoff(i, j)
            End If
            If (StrainCheckoutCoff(i, j) > StrainStatPara(i, 2)) Then
                StrainStatPara(i, 2) = StrainCheckoutCoff(i, j)
            End If
            If (RefRemainStrain(i, j) > StrainStatPara(i, 3)) Then
                StrainStatPara(i, 3) = RefRemainStrain(i, j)
            End If
        Next

        '数据写入Excel
        Cells(3, 2 * i) = Format(StrainStatPara(i, 1), "Fixed"): Cells(4, 2 * i) = Format(StrainStatPara(i, 2), "Fixed"): Cells(5, 2 * i) = Format(StrainStatPara(i, 3), "Percent")
    Next
 
End Sub
Sub AutoStrain_test()

    Dim rowCurr As Integer    '行指针
    FirstRow = 10
    
    c1 = 14: c2 = 15: c3 = 16: c4 = 17: c5 = 18: c6 = 19: c7 = 20
    
    
    rowCurr = FirstRow
    
    While Cells(rowCurr, 1) <> ""
        Cells(rowCurr, TotalDispCol) = Cells(rowCurr, TotalDispCol - 1) - Cells(rowCurr, TotalDispCol - 2)
        rowCurr = rowCurr + 1
    Wend
   
End Sub
