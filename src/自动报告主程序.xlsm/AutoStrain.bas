Attribute VB_Name = "AutoStrain"
Option Explicit

Private Const First_Row As Integer = 15    '起始数据行数

Private Const MaxElasticStrain_Row As Integer = 5
Private Const MinStrainCheckoutCoff_Row As Integer = 6
Private Const MaxStrainCheckoutCoff_Row As Integer = 7
Private Const MinRefRemainStrain_Row As Integer = 8
Private Const MaxRefRemainStrain_Row As Integer = 9

Public StrainStatPara(1 To MAX_NWC, 1 To 5)  '统计参数

Private Const e1 As Integer = 3  '初始模数（所在列，以下略）
Private Const e2 As Integer = 4  '初始温度
Private Const e3 As Integer = 5  '满载模数
Private Const e4 As Integer = 6  '满载温度
Private Const e5 As Integer = 7  '卸载模数
Private Const e6 As Integer = 8  '卸载温度

Const InstrumentType_Col As Integer = 11
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
Private Const Strain_WC_Col As Integer = 1    '应变工况所在列
Private Const StrainNode_WCCopy_Col As Integer = 25
Private Const StrainNode_NameCopy_Col As Integer = 26

Public StrainGlobalWC(1 To MAX_NWC)  As Integer  '全局工况定位数组

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


Private Const TheoryStress_Col As Integer = 9
Private Const TheoryStrain_Col As Integer = 10
Private Const Instrument_Col As Integer = 11
Private Const TheoryStrainCopy_Col As Integer = 30  '复制数据

Private Const StrainCheckoutCoff_Col As Integer = 31
Private Const RefRemainStrain_Col As Integer = 32

Public StrainUbound(1 To MAX_NWC) As Integer    '每个工况上界（下界为1）
Public strainNWCs As Integer    '应变工况数
Public StrainNPs(10) As Integer    '各个工况测点数

Public StrainChartObjArray(1 To MAX_NWC) As Shape

Public StrainGroupName(1 To MAX_NWC)  As String
Public strainResultVar(1 To MAX_NWC) As String
Public strainResult(1 To MAX_NWC) As String
Public strainSummaryVar(1 To MAX_NWC) As String
Public strainSummary(1 To MAX_NWC) As String
Public strainTbTitle(1 To MAX_NWC) As String
Public strainRawTbTitle(1 To MAX_NWC) As String
Public strainChartTitle(1 To MAX_NWC) As String

Public strainTheoryShapeTitle(1 To MAX_NWC) As String    '理论应变值
Public strainTheoryShapeTitleVar(1 To MAX_NWC) As String

Public strainRawTableBookmarks(1 To MAX_NWC) As String
Public strainTblBookmarks(1 To MAX_NWC) As String
Public strainChartBookmarks(1 To MAX_NWC) As String
Public strainTheoryShapeBookmarks(1 To MAX_NWC) As String

Public strainTbCrossRef(1 To MAX_NWC) As String    '报告交叉引用
Public strainGraphCrossRef(1 To MAX_NWC) As String


Public Sub InitStrainVar()
    Dim i As Integer
    nPN = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十")
    strainNWCs = Cells(1, 2)
    For i = 1 To strainNWCs
        StrainUbound(i) = Cells(2, 2 * i)
    Next
 
    For i = 1 To strainNWCs
        StrainGlobalWC(i) = Cells(3, 2 * i)
    Next
    
    For i = 1 To strainNWCs
        StrainGroupName(i) = Cells(4, 2 * i)
    Next
    
End Sub
'生成应变表格的行
Private Sub GenerateStrainRows_Click()

    strainNWCs = Cells(1, 2)
    
    Dim i As Integer
    For i = 0 To strainNWCs - 1    '每个工况测点数
        StrainNPs(i) = Cells(2, 2 * (i + 1))
    Next
      
    Dim colorArray(1 To 11) As Integer
    For i = 1 To 8
        colorArray(i) = i
    Next
    colorArray(9) = TheoryStress_Col
    colorArray(10) = TheoryStrain_Col
    colorArray(11) = InstrumentType_Col
    
    Dim rowCurr As Integer    '行指针
    rowCurr = First_Row
    
    Dim ds As DeformService
    Set ds = New DeformService
    ds.GenerateRows strainNWCs, StrainNPs, rowCurr, 1, 2, colorArray
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
  k = 1
  For i = 1 To 11
    dataArray(k) = i
    k = k + 1
  Next i
  For i = 12 To 19
    dataArray(k) = i    '序号从9开始
    k = k + 1
  Next i
  
  For i = 25 To 32
    dataArray(k) = i '序号从19开始
    k = k + 1
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
'Public Function GetStrain(ByVal r2, ByVal r1, ByVal t2, ByVal t1)
'    Dim G, k, c
'   G = 3.7: k = 1.8: c = 1.020019
'    GetStrain = G * c * (r2 - r1) + k * (t2 - t1)
'End Function

'计算残余应变
'deltaS:卸载状态与初始状态对比产生的应变值，totalS:总应变
'TODO:测试
'初读应变：0，满载应变：30，卸载应变：3 =>3
'初读应变：0，满载应变：30，卸载应变：-3 =>0
'初读应变：0，满载应变：-30，卸载应变：-3 =>-3
'初读应变：0，满载应变：-30，卸载应变：3 =>0
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
    
    Dim ax As ArrayService    '原as变量名与关键字as冲突，改为as
    Set ax = New ArrayService
    
    Dim ds As DeformService
    Set ds = New DeformService
    
    Dim i, j As Integer

    Dim rowCurr As Integer    '行指针
    rowCurr = First_Row
    
    For i = 1 To strainNWCs
        For j = 1 To StrainUbound(i)
        
            StrainNodeName(i, j) = Cells(rowCurr, StrainNode_Name_Col)
            
            '以下两行复制数据，方便查看
            Cells(rowCurr, StrainNode_WCCopy_Col) = Cells(rowCurr, Strain_WC_Col)
            Cells(rowCurr, StrainNode_NameCopy_Col) = Cells(rowCurr, StrainNode_Name_Col)
                     
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
            
            'TotalStrain(i, j) = GetStrain(Cells(rowCurr, e3), Cells(rowCurr, e1), Cells(rowCurr, e4), Cells(rowCurr, e2))
            TotalStrain(i, j) = ds.GetStrain(Cells(rowCurr, e3), Cells(rowCurr, e1), Cells(rowCurr, e4), Cells(rowCurr, e2), Cells(rowCurr, InstrumentType_Col))
            INTTotalStrain(i, j) = Round(TotalStrain(i, j), 0)
            
            Cells(rowCurr, TotalStrain_Col) = TotalStrain(i, j)
            Cells(rowCurr, INTTotalStrain_Col) = INTTotalStrain(i, j)
            
            RemainStrain(i, j) = GetRemainStrain(ds.GetStrain(Cells(rowCurr, e5), Cells(rowCurr, e1), Cells(rowCurr, e6), Cells(rowCurr, e2), Cells(rowCurr, InstrumentType_Col)), TotalStrain(i, j))
            'IIf(GetStrain(Cells(rowCurr, e5), Cells(rowCurr, e1), Cells(rowCurr, e6), Cells(rowCurr, e2)) >= 0 _
            ', GetStrain(Cells(rowCurr, e5), Cells(rowCurr, e1), Cells(rowCurr, e6), Cells(rowCurr, e2)), 0)
            INTRemainStrain(i, j) = Round(RemainStrain(i, j), 0)
            Cells(rowCurr, RemainStrain_Col) = RemainStrain(i, j)
            Cells(rowCurr, INTRemainStrain_Col) = INTRemainStrain(i, j)    '残余应变
            
            ElasticStrain(i, j) = TotalStrain(i, j) - RemainStrain(i, j)
            INTElasticStrain(i, j) = INTTotalStrain(i, j) - INTRemainStrain(i, j)
            Cells(rowCurr, INTElasticStrain_Col) = INTElasticStrain(i, j)    '弹性应变
             
            TheoryStrain(i, j) = Cells(rowCurr, TheoryStrain_Col)    '理论应变直接取值
            Cells(rowCurr, TheoryStrainCopy_Col) = TheoryStrain(i, j)
        
            
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
    For i = 1 To strainNWCs
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
    
    '计算各组所对应文档变量的值
    '算法：根据报告工况数量来区分
    For i = 1 To strainNWCs
         '插入对应文档变量
         strainResultVar(i) = Replace("strainResult" & CStr(i), " ", "")
         strainSummaryVar(i) = Replace("strain" & CStr(i), " ", "")
         strainTbCrossRef(i) = "strainTbCrossRef" & CStr(i)
         strainGraphCrossRef(i) = "strainGraphCrossRef" & CStr(i)
         '对应书签名
        strainRawTableBookmarks(i) = Replace("strainRawTable" & CStr(i), " ", "")
        strainTblBookmarks(i) = Replace("strainTable" & CStr(i), " ", "")
        strainChartBookmarks(i) = Replace("strainChart" & CStr(i), " ", "")
        strainTheoryShapeBookmarks(i) = Replace("strainTheoryShape" & CStr(i), " ", "")
        
         '对应文档变量
        strainTheoryShapeTitleVar(i) = Replace("strainTheoryShapeTitle" & CStr(i), " ", "")
        
        '插入结果描述以及图表概述
        If ax.CountElements(StrainGlobalWC, StrainGlobalWC(i)) = 1 Then    '工况只有1组
            strainResult(i) = "(" & CStr(i) & ")在工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & "荷载作用下，所测主梁最大弹性应变为" & Round(StrainStatPara(i, MaxElasticDeform_Index), 0) & "με，" _
                & "实测控制截面的混凝土应变值均小于理论值，校验系数在" & Format(StrainStatPara(i, MinCheckoutCoff_Index), "Fixed") & "～" & Format(StrainStatPara(i, MaxCheckoutCoff_Index), "Fixed") & "之间；" _
                & "相对残余应变在" & Format(StrainStatPara(i, MinRefRemainDeform_Index), "Percent") & "～" & Format(StrainStatPara(i, MaxRefRemainDeform_Index), "Percent") & "之间。"
            strainSummary(i) = "工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & "测试截面测点应变检测结果详见" & strainTbCrossRef(i) & "，应变实测值与理论计算值的关系曲线详见" & strainGraphCrossRef(i) & "。检测结果表明，所测主梁的应变校验系数在" _
                 & Format(StrainStatPara(i, MinCheckoutCoff_Index), "Fixed") & "～" & Format(StrainStatPara(i, MaxCheckoutCoff_Index), "Fixed") & "之间，" _
                 & "相对残余应变在" & Format(StrainStatPara(i, MinRefRemainDeform_Index), "Percent") & "～" & Format(StrainStatPara(i, MaxRefRemainDeform_Index), "Percent") & "之间。"
            strainTbTitle(i) = "工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & "应变检测结果汇总表"
            strainChartTitle(i) = "工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & "应变实测值与理论计算值的关系曲线"
            
            strainRawTbTitle(i) = "工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & "应变原始数据处理表"    '生成计算书额外使用的变量
            
            strainTheoryShapeTitle(i) = "工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & "控制截面理论应力值（单位：MPa）"
        Else
             strainResult(i) = "(" & CStr(i) & ")在工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & "荷载作用下，所测主梁" & StrainGroupName(i) & "截面最大弹性应变为" & Round(StrainStatPara(i, MaxElasticDeform_Index), 0) & "με，" _
                & "实测控制截面的混凝土应变值均小于理论值，校验系数在" & Format(StrainStatPara(i, MinCheckoutCoff_Index), "Fixed") & "～" & Format(StrainStatPara(i, MaxCheckoutCoff_Index), "Fixed") & "之间；" _
                & "相对残余应变在" & Format(StrainStatPara(i, MinRefRemainDeform_Index), "Percent") & "～" & Format(StrainStatPara(i, MaxRefRemainDeform_Index), "Percent") & "之间。"
            strainSummary(i) = "工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & "主梁" & StrainGroupName(i) & "截面测点应变检测结果详见" & strainTbCrossRef(i) & "，应变实测值与理论计算值的关系曲线详见" & strainGraphCrossRef(i) & "。检测结果表明，所测主梁的应变校验系数在" _
                 & Format(StrainStatPara(i, MinCheckoutCoff_Index), "Fixed") & "～" & Format(StrainStatPara(i, MaxCheckoutCoff_Index), "Fixed") & "之间，" _
                 & "相对残余应变在" & Format(StrainStatPara(i, MinRefRemainDeform_Index), "Percent") & "～" & Format(StrainStatPara(i, MaxRefRemainDeform_Index), "Percent") & "之间。"
            strainTbTitle(i) = "工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & StrainGroupName(i) & "截面应变检测结果汇总表"
            strainChartTitle(i) = "工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & StrainGroupName(i) & "截面应变实测值与理论计算值的关系曲线"
            
            strainRawTbTitle(i) = "工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & StrainGroupName(i) & "截面应变原始数据处理表"    '生成计算书额外使用的变量
            strainTheoryShapeTitle(i) = "工况" & CStr(nPN(StrainGlobalWC(i) - 1)) & StrainGroupName(i) & "截面理论应力值（单位：MPa）"
        End If
    Next
    
    Set ax = Nothing

    Set ds = Nothing
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
    yStep = 260    '原为220，定义为上限即可
    
    Dim chartWidth As Integer: Dim chartHeight As Integer
   
    Dim i As Integer: Dim curr As Integer: Dim currCounts As Integer   '当前作图测点数
    curr = First_Row: currCounts = 1 '至少一个点
    
    For i = 1 To strainNWCs
        currCounts = Range(Cells(curr, StrainNode_Name_Col), Cells(curr + StrainUbound(i) - 1, StrainNode_Name_Col)).Rows.Count
        chartWidth = 350: chartHeight = 220 '默认值
        If currCounts > 12 And currCounts < 15 Then
            chartWidth = 380: chartHeight = 240
        ElseIf currCounts >= 15 And currCounts < 23 Then
             chartWidth = 400: chartHeight = 250
        ElseIf currCounts >= 23 Then
             chartWidth = 450: chartHeight = 270    '高宽比与其他点数不同
        End If
        Set StrainChartObjArray(i) = Sheets(StrainSheetName).Shapes.AddChart2(332, xlLineMarkers, xPos, yPos + (i - 1) * yStep, chartWidth, chartHeight)
        StrainChartObjArray(i).Select
    
        '原代码：    'Sheets(StrainSheetName).Shapes.AddChart2(332, xlLineMarkers, xPos, yPos + (i - 1) * yStep, 350, 200).Select
        With ActiveChart
    
                'a = Range(Cells(curr, StrainNode_Name_Col), Cells(curr + StrainUbound(i) - 1, StrainNode_Name_Col)).Rows.Count
                .SetSourceData Source:=Union(Range(Cells(curr, StrainNode_Name_Col), Cells(curr + StrainUbound(i) - 1, StrainNode_Name_Col)), _
                Range(Cells(curr, INTElasticStrain_Col), Cells(curr + StrainUbound(i) - 1, INTElasticStrain_Col)), Range(Cells(curr, TheoryStrain_Col), Cells(curr + StrainUbound(i) - 1, TheoryStrain_Col)))
                
                  
                .SetElement (msoElementChartTitleNone)    '删除标题
                .SeriesCollection(1).name = "理论值"
                .SeriesCollection(2).name = "实测值"
        
                '横坐标
                '.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
                '.Axes(xlCategory, xlPrimary).AxisTitle.Text = "测点号"
                
                '纵坐标
                .SetElement msoElementPrimaryValueAxisTitleAdjacentToAxis
                
                .Axes(xlValue).HasTitle = True
                .Axes(xlValue).AxisTitle.caption = "应变（με）"
        
        End With
        
        curr = curr + StrainUbound(i)
    Next i


End Sub
