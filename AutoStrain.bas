Attribute VB_Name = "AutoStrain"
Option Explicit

Private Const First_Row As Integer = 15    '起始数据行数

Private Const MaxElasticStrain_Row As Integer = 4
Private Const MinStrainCheckoutCoff_Row As Integer = 5
Private Const MaxStrainCheckoutCoff_Row As Integer = 6
Private Const MinRefRemainStrain_Row As Integer = 7
Private Const MaxRefRemainStrain_Row As Integer = 8

Public StrainStatPara(1 To MAX_NWC, 1 To 5)  '统计参数

Private Const e1 As Integer = 3  '初始模数（所在列，以下略）
Private Const e2 As Integer = 4  '初始温度
Private Const e3 As Integer = 5  '满载模数
Private Const e4 As Integer = 6  '满载温度
Private Const e5 As Integer = 7  '卸载模数
Private Const e6 As Integer = 8  '卸载温度

Const FullLoadStrainR_Col As Integer = 14
Const FullLoadStrainT_Col As Integer = 15
Const UnLoadStrainR_Col As Integer = 17
Const UnLoadStrainT_Col As Integer = 18

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


Const StrainNode_Name_Col As Integer = 2  '测点编号所在列

Public StrainGlobalWC(1 To MAX_NWC)    '全局工况定位数组

Public StrainNodeName(1 To MAX_NWC, 1 To MAX_NPS) As String  '各个工况测点名称

Public InitStrainR0(1 To MAX_NWC, 1 To MAX_NPS)
Public InitStrainT0(1 To MAX_NWC, 1 To MAX_NPS)
Public FullLoadStrainR0(1 To MAX_NWC, 1 To MAX_NPS)
Public FullLoadStrainT0(1 To MAX_NWC, 1 To MAX_NPS)
Public UnLoadStrainR0(1 To MAX_NWC, 1 To MAX_NPS)
Public UnLoadStrainT0(1 To MAX_NWC, 1 To MAX_NPS)

Public FullLoadStrainR(1 To MAX_NWC, 1 To MAX_NPS)
Public FullLoadStrainT(1 To MAX_NWC, 1 To MAX_NPS)
Public UnLoadStrainR(1 To MAX_NWC, 1 To MAX_NPS)
Public UnLoadStrainT(1 To MAX_NWC, 1 To MAX_NPS)

Public TotalStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double    '总应变
Public RemainStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double    '残余应变
Public ElasticStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double
Public TheoryStress(1 To MAX_NWC, 1 To MAX_NPS) As Double
Public TheoryStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double
Public StrainCheckoutCoff(1 To MAX_NWC, 1 To MAX_NPS) As Double
Public RefRemainStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double

Public INTTotalStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double '满载总应变（取整后，下同）
Public INTElasticStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double
Public INTRemainStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double
Public INTDivStrainCheckoutCoff(1 To MAX_NWC, 1 To MAX_NPS) As Double   '整数应变值运算所得结果（下同）
Public INTDivRefRemainStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double

Private Const INTTotalStrain_Col As Integer = 27
Private Const INTRemainStrain_Col As Integer = 29
Private Const INTElasticStrain_Col As Integer = 28

Private Const TotalStrain_Col As Integer = 16
Private Const RemainStrain_Col As Integer = 19

Private Const TheoryStress_Col As Integer = 20

Private Const TheoryStrain_Col As Integer = 30
Private Const StrainCheckoutCoff_Col As Integer = 31
Private Const RefRemainStrain_Col As Integer = 32

Public StrainUbound(1 To MAX_NWC) As Integer    '每个工况上界（下界为1）
Public StrainNWCs As Integer    '应变工况数
Public StrainNPs(10) As Integer    '各个工况测点数
Public Sub InitStrainVar()
    Dim i As Integer
    
    StrainNWCs = Cells(1, 2)
    For i = 1 To StrainNWCs
        StrainUbound(i) = Cells(2, 2 * i)
    Next
 
    For i = 1 To StrainNWCs
        StrainGlobalWC(i) = Cells(3, 2 * i)
    Next
End Sub
'生成应变表格的行
Private Sub GenerateStrainRows_Click()

    StrainNWCs = Cells(1, 2)
    
    Dim i As Integer
    For i = 0 To StrainNWCs - 1    '每个工况测点数
        StrainNPs(i) = Cells(2, 2 * (i + 1))
    Next
      
    Dim colorArray(1 To 9) As Integer
    For i = 1 To 8
        colorArray(i) = i
    Next
    colorArray(9) = TheoryStrain_Col
    
    Dim rowCurr As Integer    '行指针
    rowCurr = First_Row
    
    Dim ds As DeformService
    Set ds = New DeformService
    ds.GenerateRows StrainNWCs, StrainNPs, rowCurr, 1, colorArray
    Set ds = Nothing
 
End Sub

'清空数据
Public Sub StrainDataClear()
  If (MsgBox("清空输入数据不可撤销，你确定要清空吗？", vbYesNo + vbExclamation, "该操作不可撤销") = vbNo) Then
    Exit Sub
  End If
  
  Dim i, j, k As Integer
  Dim rowCurr As Integer    '行指针
  rowCurr = First_Row
  
  'TODO:数组初始化更灵活
  Dim dataArray(1 To 100) As Integer    '待清空的列
  For i = 1 To 8
    dataArray(i) = i
  Next i
  For i = 12 To 21
    dataArray(i - 3) = i    '序号从9开始
  Next i
  
  For i = 25 To 32
    dataArray(i - (25 - 18 - 1)) = i '序号从19开始
  Next i
  
  
  '清空表格数据
  While Cells(rowCurr, 1) <> ""    '第一个单元格数据作为判断依据
    For i = 1 To UBound(dataArray)
        If dataArray(i) = 0 Then Exit For
        Cells(rowCurr, dataArray(i)) = ""
        Cells(rowCurr, dataArray(i)).Interior.Color = RGB(255, 255, 255) ' RGB(0, 176, 80)
    Next
    rowCurr = rowCurr + 1
  Wend

  '清空统计数据
    For i = 1 To MAX_NWC
        Cells(MaxElasticStrain_Row, 2 * i) = ""
        Cells(MinStrainCheckoutCoff_Row, 2 * i) = ""
        Cells(MaxStrainCheckoutCoff_Row, 2 * i) = ""
        Cells(MinRefRemainStrain_Row, 2 * i) = ""
        Cells(MaxRefRemainStrain_Row, 2 * i) = ""
    Next
End Sub

'计算应变
'r2:变化后模数，r1:变化前模数，t2:变化后温度，t1：变化前温度
Public Function GetStrain(ByVal r2, ByVal r1, ByVal t2, ByVal t1)
    Dim G, k, c
    G = 3.7: k = 1.8: c = 1.020019
    GetStrain = G * c * (r2 - r1) + k * (t2 - t1)
End Function

'计算残余应变
'deltaS:卸载状态与初始状态对比产生的应变值，totalS:总应变
Public Function GetRemainStrain(ByVal deltaS, ByVal totalS)
    If totalS >= 0 Then
        If deltaS >= 0 Then
            GetRemainStrain = deltaS
        Else
            GetRemainStrain = 0
        End If
    ElseIf totalS < 0 Then
        If deltaS <= 0 Then
            GetRemainStrain = deltaS
        Else
           GetRemainStrain = 0
        End If
    End If
End Function

'自动计算应变
Public Sub AutoStrain_Click()

    InitStrainVar
    
    Dim i, j As Integer

    Dim rowCurr As Integer    '行指针
    rowCurr = First_Row
    
    For i = 1 To StrainNWCs
        For j = 1 To StrainUbound(i)
        
            StrainNodeName(i, j) = Cells(rowCurr, StrainNode_Name_Col)
            TheoryStress(i, j) = Cells(rowCurr, TheoryStress_Col)
            InitStrainR0(i, j) = Cells(rowCurr, e1)
            InitStrainT0(i, j) = Cells(rowCurr, e2)
            FullLoadStrainR0(i, j) = Cells(rowCurr, e3)
            FullLoadStrainT0(i, j) = Cells(rowCurr, e4)
            UnLoadStrainR0(i, j) = Cells(rowCurr, e5)
            UnLoadStrainT0(i, j) = Cells(rowCurr, e6)
                     
            FullLoadStrainR(i, j) = Cells(rowCurr, e3) - Cells(rowCurr, e1)
            Cells(rowCurr, FullLoadStrainR_Col) = FullLoadStrainR(i, j)
            
            FullLoadStrainT(i, j) = Cells(rowCurr, e4) - Cells(rowCurr, e2)
            Cells(rowCurr, FullLoadStrainT_Col) = FullLoadStrainT(i, j)
            
            UnLoadStrainR(i, j) = Cells(rowCurr, e5) - Cells(rowCurr, e1)
            Cells(rowCurr, UnLoadStrainR_Col) = UnLoadStrainR(i, j)
            
            UnLoadStrainT(i, j) = Cells(rowCurr, e6) - Cells(rowCurr, e2)
            Cells(rowCurr, UnLoadStrainT_Col) = UnLoadStrainT(i, j)
            
            TotalStrain(i, j) = GetStrain(Cells(rowCurr, e3), Cells(rowCurr, e1), Cells(rowCurr, e4), Cells(rowCurr, e2))
            INTTotalStrain(i, j) = Round(TotalStrain(i, j), 0)
            
            Cells(rowCurr, TotalStrain_Col) = TotalStrain(i, j)
            Cells(rowCurr, INTTotalStrain_Col) = INTTotalStrain(i, j)
            
            RemainStrain(i, j) = GetRemainStrain(GetStrain(Cells(rowCurr, e5), Cells(rowCurr, e1), Cells(rowCurr, e6), Cells(rowCurr, e2)), TotalStrain(i, j))
            'IIf(GetStrain(Cells(rowCurr, e5), Cells(rowCurr, e1), Cells(rowCurr, e6), Cells(rowCurr, e2)) >= 0 _
            ', GetStrain(Cells(rowCurr, e5), Cells(rowCurr, e1), Cells(rowCurr, e6), Cells(rowCurr, e2)), 0)
            INTRemainStrain(i, j) = Round(RemainStrain(i, j), 0)
            Cells(rowCurr, RemainStrain_Col) = RemainStrain(i, j)
            Cells(rowCurr, INTRemainStrain_Col) = INTRemainStrain(i, j)    '残余应变
            
            ElasticStrain(i, j) = TotalStrain(i, j) - RemainStrain(i, j)
            INTElasticStrain(i, j) = INTTotalStrain(i, j) - INTRemainStrain(i, j)
            Cells(rowCurr, INTElasticStrain_Col) = INTElasticStrain(i, j)    '弹性应变
             
            TheoryStrain(i, j) = Cells(rowCurr, TheoryStrain_Col)    '理论应变直接取值
            
            StrainCheckoutCoff(i, j) = ElasticStrain(i, j) / TheoryStrain(i, j)
            INTDivStrainCheckoutCoff(i, j) = INTElasticStrain(i, j) / TheoryStrain(i, j)
            Cells(rowCurr, StrainCheckoutCoff_Col) = INTDivStrainCheckoutCoff(i, j)    '校验系数
            
            If TotalStrain(i, j) = 0 Then
                RefRemainStrain(i, j) = 0
            Else
                RefRemainStrain(i, j) = RemainStrain(i, j) / TotalStrain(i, j)
            End If
            
            If INTTotalStrain(i, j) = 0 Then
                INTDivRefRemainStrain(i, j) = 0
            Else
                INTDivRefRemainStrain(i, j) = INTRemainStrain(i, j) / INTTotalStrain(i, j)
            End If
            Cells(rowCurr, RefRemainStrain_Col) = INTDivRefRemainStrain(i, j)    '相对残余变形
            
            rowCurr = rowCurr + 1
        Next
    Next
    
    '注意！总应变、弹性应变、残余应变数值均为取整后计算的！
    '计算各个工况最小/大校验系数，最大相对残余应变
    For i = 1 To StrainNWCs
        StrainStatPara(i, MaxElasticDeform_Index) = INTElasticStrain(i, 1): StrainStatPara(i, MinCheckoutCoff_Index) = INTDivStrainCheckoutCoff(i, 1): StrainStatPara(i, MaxCheckoutCoff_Index) = INTDivStrainCheckoutCoff(i, 1)
        StrainStatPara(i, MinRefRemainDeform_Index) = INTDivRefRemainStrain(i, 1): StrainStatPara(i, MaxRefRemainDeform_Index) = INTDivRefRemainStrain(i, 1)
        For j = 1 To StrainUbound(i)
            If (INTElasticStrain(i, j) > StrainStatPara(i, MaxElasticDeform_Index)) Then
                StrainStatPara(i, 1) = INTElasticStrain(i, j)
            End If
            If (INTDivStrainCheckoutCoff(i, j) < StrainStatPara(i, MinCheckoutCoff_Index)) Then
                StrainStatPara(i, 2) = INTDivStrainCheckoutCoff(i, j)
            End If
            If (INTDivStrainCheckoutCoff(i, j) > StrainStatPara(i, MaxCheckoutCoff_Index)) Then
                StrainStatPara(i, 3) = INTDivStrainCheckoutCoff(i, j)
            End If
            If (INTDivRefRemainStrain(i, j) < StrainStatPara(i, MinRefRemainDeform_Index)) Then
                StrainStatPara(i, 4) = INTDivRefRemainStrain(i, j)
            End If
            If (INTDivRefRemainStrain(i, j) > StrainStatPara(i, MaxRefRemainDeform_Index)) Then
                StrainStatPara(i, 5) = INTDivRefRemainStrain(i, j)
            End If
        Next

        '数据写入Excel
        Cells(MaxElasticStrain_Row, 2 * i) = Format(StrainStatPara(i, MaxElasticDeform_Index), "Fixed"): Cells(MinStrainCheckoutCoff_Row, 2 * i) = Format(StrainStatPara(i, MinCheckoutCoff_Index), "Fixed")
        Cells(MaxStrainCheckoutCoff_Row, 2 * i) = Format(StrainStatPara(i, MaxCheckoutCoff_Index), "Fixed"): Cells(MinRefRemainStrain_Row, 2 * i) = Format(StrainStatPara(i, MinRefRemainDeform_Index), "Percent")
        Cells(MaxRefRemainStrain_Row, 2 * i) = Format(StrainStatPara(i, MaxRefRemainDeform_Index), "Percent")
    Next
 
End Sub

'自动作图（需要先计算）
Public Sub AutoStrainGraph()
    'https://docs.microsoft.com/zh-CN/office/vba/api/Excel.shapes.addchart2
    'AddChart2(样式， XlChartType， 左， 顶部， 宽度， 高度， NewLayout)
    Dim StrainSheetName As String
    StrainSheetName = "应变"
    
    Dim xPos, yPos As Integer  '定义第一张图x,y位置
    xPos = 800: yPos = 150
    Dim yStep As Integer    'y方向每个图表占用空间
    yStep = 220
      
   
    Dim i As Integer
    Dim curr As Integer
    curr = First_Row
    
    For i = 1 To StrainNWCs
        Sheets(StrainSheetName).Shapes.AddChart2(332, xlLineMarkersStacked, xPos, yPos + (i - 1) * yStep, 350, 200).Select
        With ActiveChart
    
                .SetSourceData Source:=Union(Range(Cells(curr, StrainNode_Name_Col), Cells(curr + StrainUbound(i) - 1, StrainNode_Name_Col)), _
                Range(Cells(curr, INTElasticStrain_Col), Cells(curr + StrainUbound(i) - 1, INTElasticStrain_Col)), Range(Cells(curr, TheoryStrain_Col), Cells(curr + StrainUbound(i) - 1, TheoryStrain_Col)))
                
                .SetElement (msoElementChartTitleNone)    '删除标题
                .SeriesCollection(1).Name = "实测值"
                .SeriesCollection(2).Name = "理论值"
        
                '横坐标
                .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
                .Axes(xlCategory, xlPrimary).AxisTitle.Text = "测点号"
                '纵坐标
                .SetElement msoElementPrimaryValueAxisTitleAdjacentToAxis
                
                .Axes(xlValue).HasTitle = True
                .Axes(xlValue).AxisTitle.Caption = "应变（με）"
        
        End With
        
        curr = curr + DispUbound(i)
    Next i


End Sub
