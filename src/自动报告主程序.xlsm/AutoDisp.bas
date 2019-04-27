Attribute VB_Name = "AutoDisp"
Option Explicit
Const WC_Col As Integer = 1     '工况所在列数
Public Const MAX_NWC As Integer = 10     '最大工况数
Public Const MAX_NPS As Integer = 100     '每个工况最大测点数

Private Const First_Row As Integer = 13     '起始数据行数

Private Const MaxElasticDisp_Row As Integer = 5
Private Const MinCheckoutCoff_Row As Integer = 6
Private Const MaxCheckoutCoff_Row As Integer = 7
Private Const MinRefRemainDisp_Row As Integer = 8
Private Const MaxRefRemainDisp_Row As Integer = 9

Public StatPara(1 To MAX_NWC, 1 To 5)  '统计参数


Const Node_Name_Col As Integer = 2  '测点编号所在列（下同）

Const InitDisp_Col As Integer = 3  '初始读数
Const FullLoadDisp_Col As Integer = 4  '满载读数
Const UnLoadDisp_Col As Integer = 5  '卸载读数
Const TheoryDisp_Col As Integer = 6  '理论位移所在列

Const TotalDispCol As Integer = 16   '总变形所在列
Const ElasticCol As Integer = 17   '弹性变形所在列
Const RemainDispCol As Integer = 18  '残余变形所在列
Private Const TheoryDispCopy_Col As Integer = 19  '复制数据
Const CheckoutCoffCol As Integer = 20   '校验系数所在列
Const RefRemainDispCol As Integer = 21  '相对残余变形所在列

Const DeltaCol As Integer = 22  '增量所在列

Const GroupNameCol As Integer = 13   '分组名称所在列

Public NWCs As Integer    '（挠度）工况数

Public GlobalWC(1 To MAX_NWC)  As Integer   '全局工况定位数组
Public GroupName(1 To MAX_NWC)  As String '每个组的名称

'Obsolete code:
'Public GroupName(1 To MAX_NWC, 1 To MAX_NPS)  As String '各个工况各个测点所属分组
'如GroupName(i,j)表示第i个工况第j个测点所属分组名称

Public InitDisp(1 To MAX_NWC, 1 To MAX_NPS)  '初始读数
Public FullLoadDisp(1 To MAX_NWC, 1 To MAX_NPS)  '满载读数
Public UnLoadDisp(1 To MAX_NWC, 1 To MAX_NPS)  '卸载读数

Public TotalDisp(1 To MAX_NWC, 1 To 100)   'TotalDisp(i,j)表示第i个工况，第j个测点总变形
Public NodeName(1 To MAX_NWC, 1 To 100) As String  '各个工况测点名称
Public Delta(1 To MAX_NWC, 1 To 100)    '增量，备用变量
Public RemainDisp(1 To MAX_NWC, 1 To 100)
Public ElasticDisp(1 To MAX_NWC, 1 To 100)
Public TheoryDisp(1 To MAX_NWC, 1 To 100)
Public CheckoutCoff(1 To MAX_NWC, 1 To 100)
Public RefRemainDisp(1 To MAX_NWC, 1 To 100)
Public DispUbound(1 To MAX_NWC) As Integer    '每个工况上界（下界为1）

Public DispChartObjArray(1 To MAX_NWC) As Shape    'as Object    '存储各图表指针

Public dispResultVar(1 To MAX_NWC) As String    '对应文档变量名
Public dispResult(1 To MAX_NWC) As String    '和word中响应DocVariable对应
Public dispSummaryVar(1 To MAX_NWC) As String    '对应文档变量名
Public dispSummary(1 To MAX_NWC) As String    '和word中响应DocVariable对应
Public dispTbTitle(1 To MAX_NWC) As String
Public dispTbCrossRef(1 To MAX_NWC) As String    '报告交叉引用
Public dispGraphCrossRef(1 To MAX_NWC) As String

Public dispRawTbTitle(1 To MAX_NWC) As String
Public dispChartTitle(1 To MAX_NWC) As String

Public dispTheoryShapeTitle(1 To MAX_NWC) As String    '理论挠度值
Public dispTheoryShapeTitleVar(1 To MAX_NWC) As String

Public dispRawTblBookmarks(1 To MAX_NWC) As String
Public dispTblBookmarks(1 To MAX_NWC) As String
Public dispChartBookmarks(1 To MAX_NWC) As String
Public dispTheoryShapeBookmarks(1 To MAX_NWC) As String


'清空数据
Public Sub DispDataClear()
  If (MsgBox("清空输入数据不可撤销，你确定要清空吗？", vbYesNo + vbExclamation, "该操作不可撤销") = vbNo) Then
    Exit Sub
  End If
  
  Dim i, j As Integer
  Dim rowCurr As Integer    '行指针
  rowCurr = First_Row
  
  '清空表格数据
  While Cells(rowCurr, 1) <> ""    '第一个单元格数据作为判断依据
    For i = 1 To TheoryDisp_Col
        Cells(rowCurr, i) = ""
        Cells(rowCurr, i).Interior.Color = RGB(255, 255, 255) ' RGB(0, 176, 80)
    Next
    rowCurr = rowCurr + 1
  Wend
  
  '清空统计数据
    For i = 1 To MAX_NWC
        Cells(MaxElasticDisp_Row, 2 * i) = ""
        Cells(MinCheckoutCoff_Row, 2 * i) = ""
        Cells(MaxCheckoutCoff_Row, 2 * i) = ""
        Cells(MinRefRemainDisp_Row, 2 * i) = ""
        Cells(MaxRefRemainDisp_Row, 2 * i) = ""
    Next
End Sub
'''初始化全局变量
Public Sub InitVar()

    nPN = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十")
    NWCs = Cells(1, 2)
    Dim i As Integer
    For i = 1 To NWCs
        DispUbound(i) = Cells(2, 2 * i)
    Next

    For i = 1 To NWCs
        GlobalWC(i) = Cells(3, 2 * i)
    Next
    
   For i = 1 To NWCs
        GroupName(i) = Cells(4, 2 * i)
    Next
            
End Sub


Public Sub AutoDisp_Click()

    InitVar
    
    Dim ds As DeformService
    Set ds = New DeformService
    
    Dim gs As GroupService
    Set gs = New GroupService
    
    Dim ax As ArrayService    '原as变量名与关键字as冲突，改为as
    Set ax = New ArrayService
    
    Dim rowCurr As Integer    '行指针
    Dim i, j As Integer
    rowCurr = First_Row
    For i = 1 To NWCs
        For j = 1 To DispUbound(i)
        
            NodeName(i, j) = Cells(rowCurr, Node_Name_Col)
            
            TheoryDisp(i, j) = Cells(rowCurr, TheoryDisp_Col)
            
            TheoryDisp(i, j) = Round(TheoryDisp(i, j), 2)    '取2位小数
            Cells(rowCurr, TheoryDispCopy_Col) = TheoryDisp(i, j)
            
            InitDisp(i, j) = Cells(rowCurr, InitDisp_Col)
            FullLoadDisp(i, j) = Cells(rowCurr, FullLoadDisp_Col)
            UnLoadDisp(i, j) = Cells(rowCurr, UnLoadDisp_Col)
            
            
            TotalDisp(i, j) = FullLoadDisp(i, j) - InitDisp(i, j) '总变形
            Cells(rowCurr, TotalDispCol) = TotalDisp(i, j)
            
            Cells(rowCurr, DeltaCol) = UnLoadDisp(i, j) - FullLoadDisp(i, j) '增量
            '增量不存储
            
            '算法：卸载与满载读数差值>=0，取卸载与满载读数差值，否则取0
            'RemainDisp(i, j) = IIf(UnLoadDisp(i, j) - InitDisp(i, j) >= 0, UnLoadDisp(i, j) - InitDisp(i, j), 0)
            RemainDisp(i, j) = ds.GetRemainDisp(InitDisp(i, j), FullLoadDisp(i, j), UnLoadDisp(i, j))
            Cells(rowCurr, RemainDispCol) = RemainDisp(i, j)    '残余变形
            
            ElasticDisp(i, j) = TotalDisp(i, j) - RemainDisp(i, j)
            Cells(rowCurr, ElasticCol) = ElasticDisp(i, j)    '弹性变形
             
            CheckoutCoff(i, j) = ElasticDisp(i, j) / TheoryDisp(i, j)
            Cells(rowCurr, CheckoutCoffCol) = CheckoutCoff(i, j)    '校验系数
             
            RefRemainDisp(i, j) = RemainDisp(i, j) / TotalDisp(i, j)
            Cells(rowCurr, RefRemainDispCol) = RefRemainDisp(i, j)    '相对残余变形
            
            rowCurr = rowCurr + 1
        Next
    Next
    
    '计算各个工况统计参数
    For i = 1 To NWCs
        StatPara(i, MaxElasticDeform_Index) = ElasticDisp(i, 1): StatPara(i, MinCheckoutCoff_Index) = CheckoutCoff(i, 1): StatPara(i, MaxCheckoutCoff_Index) = CheckoutCoff(i, 1)
        StatPara(i, MinRefRemainDeform_Index) = RefRemainDisp(i, 1): StatPara(i, MaxRefRemainDeform_Index) = RefRemainDisp(i, 1)
        For j = 1 To DispUbound(i)
            If (ElasticDisp(i, j) > StatPara(i, MaxElasticDeform_Index)) Then
                StatPara(i, 1) = ElasticDisp(i, j)
            End If
            If (CheckoutCoff(i, j) < StatPara(i, MinCheckoutCoff_Index)) Then
                StatPara(i, 2) = CheckoutCoff(i, j)
            End If
            If (CheckoutCoff(i, j) > StatPara(i, MaxCheckoutCoff_Index)) Then
                StatPara(i, 3) = CheckoutCoff(i, j)
            End If
            If (RefRemainDisp(i, j) < StatPara(i, MinRefRemainDeform_Index)) Then
                StatPara(i, 4) = RefRemainDisp(i, j)
            End If
            If (RefRemainDisp(i, j) > StatPara(i, MaxRefRemainDeform_Index)) Then
                StatPara(i, 5) = RefRemainDisp(i, j)
            End If
        Next
        
        '数据写入Excel
        Cells(MaxElasticDisp_Row, 2 * i) = Format(StatPara(i, 1), "Fixed"): Cells(MinCheckoutCoff_Row, 2 * i) = Format(StatPara(i, 2), "Fixed")
        Cells(MaxCheckoutCoff_Row, 2 * i) = Format(StatPara(i, MaxCheckoutCoff_Index), "Fixed"): Cells(MinRefRemainDisp_Row, 2 * i) = Format(StatPara(i, MinRefRemainDeform_Index), "Percent")
        Cells(MaxRefRemainDisp_Row, 2 * i) = Format(StatPara(i, MaxRefRemainDeform_Index), "Percent")
    Next
    
    '计算各组所对应文档变量的值
    '算法：根据报告工况数量来区分
    For i = 1 To NWCs
         '插入对应文档变量
         dispResultVar(i) = Replace("dispResult" & CStr(i), " ", "")
         dispSummaryVar(i) = Replace("dispSummary" & CStr(i), " ", "")
         dispTbCrossRef(i) = "dispTbCrossRef" & CStr(i)
         dispGraphCrossRef(i) = "dispGraphCrossRef" & CStr(i)
         
         '对应书签名
        dispRawTblBookmarks(i) = Replace("dispRawTb" & CStr(i), " ", "")
        dispTblBookmarks(i) = Replace("dispTable" & CStr(i), " ", "")
        dispChartBookmarks(i) = Replace("dispChart" & CStr(i), " ", "")
        dispTheoryShapeBookmarks(i) = Replace("dispTheoryShape" & CStr(i), " ", "")
        
        '对应文档变量
        dispTheoryShapeTitleVar(i) = Replace("dispTheoryShapeTitle" & CStr(i), " ", "")
        
        '插入结果描述以及图表概述
        If ax.CountElements(GlobalWC, GlobalWC(i)) = 1 Then    '工况只有1组
            dispResult(i) = "(" & CStr(i) & ")在工况" & CStr(nPN(GlobalWC(i) - 1)) & "荷载作用下，主梁最大实测弹性挠度值为" & Format(StatPara(i, MaxElasticDeform_Index), "Fixed") & "mm，" _
                & "实测控制截面的挠度值均小于理论值，校验系数在" & Format(StatPara(i, MinCheckoutCoff_Index), "Fixed") & "～" & Format(StatPara(i, MaxCheckoutCoff_Index), "Fixed") & "之间；" _
                & "相对残余变形在" & Format(StatPara(i, MinRefRemainDeform_Index), "Percent") & "～" & Format(StatPara(i, MaxRefRemainDeform_Index), "Percent") & "之间。"
            dispSummary(i) = "工况" & CStr(nPN(GlobalWC(i) - 1)) & "主梁挠度检测结果详见" & dispTbCrossRef(i) & "，挠度实测值与理论计算值的关系曲线详见" & dispGraphCrossRef(i) & "。检测结果表明，所测主梁的挠度校验系数在" & Format(StatPara(i, MinCheckoutCoff_Index), "Fixed") & "～" & Format(StatPara(i, MaxCheckoutCoff_Index), "Fixed") & "之间，" _
                & "相对残余变形在" & Format(StatPara(i, MinRefRemainDeform_Index), "Percent") & "～" & Format(StatPara(i, MaxRefRemainDeform_Index), "Percent") & "之间。"
            dispTbTitle(i) = "工况" & CStr(nPN(GlobalWC(i) - 1)) & "挠度检测结果汇总表"
            dispChartTitle(i) = "工况" & CStr(nPN(GlobalWC(i) - 1)) & "挠度实测值与理论计算值的关系曲线"
            
            dispTheoryShapeTitle(i) = "工况" & CStr(nPN(GlobalWC(i) - 1)) & "控制截面理论挠度值（单位：mm）"
            dispRawTbTitle(i) = "工况" & CStr(nPN(GlobalWC(i) - 1)) & "挠度原始数据处理表"    '计算书额外使用的变量
        Else
             dispResult(i) = "(" & CStr(i) & ")在工况" & CStr(nPN(GlobalWC(i) - 1)) & "荷载作用下，主梁" & GroupName(i) & "截面最大实测弹性挠度值为" & Format(StatPara(i, MaxElasticDeform_Index), "Fixed") & "mm，" _
                & "实测控制截面的挠度值均小于理论值，校验系数在" & Format(StatPara(i, MinCheckoutCoff_Index), "Fixed") & "～" & Format(StatPara(i, MaxCheckoutCoff_Index), "Fixed") & "之间；" _
                & "相对残余变形在" & Format(StatPara(i, MinRefRemainDeform_Index), "Percent") & "～" & Format(StatPara(i, MaxRefRemainDeform_Index), "Percent") & "之间。"
            dispSummary(i) = "工况" & CStr(nPN(GlobalWC(i) - 1)) & "主梁" & GroupName(i) & "截面挠度检测结果详见" & dispTbCrossRef(i) & "，挠度实测值与理论计算值的关系曲线详见" & dispGraphCrossRef(i) & "。检测结果表明，所测主梁的挠度校验系数在" & Format(StatPara(i, MinCheckoutCoff_Index), "Fixed") & "～" & Format(StatPara(i, MaxCheckoutCoff_Index), "Fixed") & "之间，" _
                & "相对残余变形在" & Format(StatPara(i, MinRefRemainDeform_Index), "Percent") & "～" & Format(StatPara(i, MaxRefRemainDeform_Index), "Percent") & "之间。"
            dispTbTitle(i) = "工况" & CStr(nPN(GlobalWC(i) - 1)) & GroupName(i) & "截面挠度检测结果汇总表"
            dispChartTitle(i) = "工况" & CStr(nPN(GlobalWC(i) - 1)) & GroupName(i) & "截面挠度实测值与理论计算值的关系曲线"
            
            dispTheoryShapeTitle(i) = "工况" & CStr(nPN(GlobalWC(i) - 1)) & GroupName(i) & "截面理论挠度值（单位：mm）"
            dispRawTbTitle(i) = "工况" & CStr(nPN(GlobalWC(i) - 1)) & GroupName(i) & "截面挠度原始数据处理表"    '计算书额外使用的变量
        End If

    Next
    Set ds = Nothing
    Set gs = Nothing
    Set ax = Nothing
End Sub

Private Sub GenerateRows_Click()
    Dim NWCs As Integer    '工况数
    Dim nPs(10) As Integer    '各个工况测点数
    Dim nPN     '各个工况对应中文名称
    nPN = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十")
    Dim i, j As Integer
    NWCs = Cells(1, 2)
    For i = 0 To NWCs - 1
        nPs(i) = Cells(2, 2 * (i + 1))
        'Debug.Print nPs(i)
    Next
    'Debug.Print nWCs
       
    Dim rowCurr As Integer    '行指针
    rowCurr = First_Row
    
    For i = 0 To NWCs - 1    '遍历工况
        For j = 1 To nPs(i)    '遍历各个工况的测点
            Cells(rowCurr, WC_Col) = nPN(i)
            
            '设置必填项的背景色
            Cells(rowCurr, WC_Col).Interior.Color = RGB(0, 176, 80)
            
            Cells(rowCurr, Node_Name_Col).Interior.Color = RGB(0, 176, 80)
            Cells(rowCurr, Node_Name_Col).NumberFormatLocal = "0.00"
            
            Cells(rowCurr, InitDisp_Col).Interior.Color = RGB(0, 176, 80)
            Cells(rowCurr, FullLoadDisp_Col).Interior.Color = RGB(0, 176, 80)
            Cells(rowCurr, UnLoadDisp_Col).Interior.Color = RGB(0, 176, 80)
            Cells(rowCurr, TheoryDisp_Col).Interior.Color = RGB(0, 176, 80)
            
            rowCurr = rowCurr + 1
        Next
    Next
    
    'TODO:添加背景色
    
End Sub

'自动作图（需要先计算）
Private Sub AutoGraph_Click()
    'https://docs.microsoft.com/zh-CN/office/vba/api/Excel.shapes.addchart2
    'AddChart2(样式， XlChartType， 左， 顶部， 宽度， 高度， NewLayout)
    Dim DispSheetName As String
    DispSheetName = "挠度"
    
    Dim i As Integer
    Dim xPos, yPos As Integer  '定义第一张图x,y位置
    xPos = 400: yPos = 150
    Dim yStep As Integer    'y方向每个图表占用空间
    yStep = 260    '原为220，定义为上限即可
    
    Dim chartWidth As Integer: Dim chartHeight As Integer
   
    Dim curr As Integer: Dim currCounts As Integer   '当前作图测点数
    curr = First_Row: currCounts = 1 '至少一个点
    
    Dim plot As Excel.Shape

    'Set plot = ws.Shapes.AddChart
    
    For i = 1 To NWCs
'        MS EXCEL VBA 代码1
'        Set plot = Sheets(DispSheetName).Shapes.AddChart2(332, xlLineMarkersStacked, xPos, yPos + (i - 1) * yStep)
'        plot.Chart.SetSourceData Source:=Union(Range(Cells(curr, Node_Name_Col), Cells(curr + DispUbound(i) - 1, Node_Name_Col)), _
'        Range(Cells(curr, ElasticCol), Cells(curr + DispUbound(i) - 1, ElasticCol)), Range(Cells(curr, TheoryDisp_Col), Cells(curr + DispUbound(i) - 1, TheoryDisp_Col)))
'
'        plot.Chart.SetElement (msoElementChartTitleNone)    '删除标题
'
'        plot.Chart.SeriesCollection(1).Name = CStr(Cells(11, 9))
'        plot.Chart.SeriesCollection(2).Name = CStr(Cells(11, 10))
'
'        '横坐标
'        plot.Chart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
'        plot.Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "测点号"
'        '纵坐标
'        plot.Chart.SetElement msoElementPrimaryValueAxisTitleBelowAxis
'        plot.Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "挠度值（mm）"

    currCounts = Range(Cells(curr, Node_Name_Col), Cells(curr + DispUbound(i) - 1, Node_Name_Col)).Rows.Count
    chartWidth = 350: chartHeight = 220 '默认值
    If currCounts > 12 And currCounts < 15 Then
        chartWidth = 380: chartHeight = 240
    ElseIf currCounts >= 15 And currCounts < 23 Then
         chartWidth = 410: chartHeight = 250
    ElseIf currCounts >= 23 Then
         chartWidth = 450: chartHeight = 270    '高宽比与其他点数不同
    End If
    Set DispChartObjArray(i) = Sheets(DispSheetName).Shapes.AddChart2(332, xlLineMarkers, xPos, yPos + (i - 1) * yStep, chartWidth, chartHeight)
    DispChartObjArray(i).Select
    '源代码：'Sheets(DispSheetName).Shapes.AddChart2(332, xlLineMarkers, xPos, yPos + (i - 1) * yStep, 350, 200).Select
    With ActiveChart
            
            .SetSourceData Source:=Union(Range(Cells(curr, Node_Name_Col), Cells(curr + DispUbound(i) - 1, Node_Name_Col)), _
            Range(Cells(curr, ElasticCol), Cells(curr + DispUbound(i) - 1, ElasticCol)), Range(Cells(curr, TheoryDisp_Col), Cells(curr + DispUbound(i) - 1, TheoryDisp_Col)))
            
            
            .SetElement (msoElementChartTitleNone)    '删除标题
            .SeriesCollection(1).name = "理论值" 'CStr(Cells(11, 9))
            .SeriesCollection(2).name = "实测值" ' CStr(Cells(11, 10))
            .Axes(xlValue).TickLabels.NumberFormatLocal = "#,##0.00_ "

            '横坐标
             '.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
             '.Axes(xlCategory, xlPrimary).AxisTitle.Text = "测点号"
                
            '纵坐标
            'ActiveChart.SetElement msoElementPrimaryValueAxisTitleBelowAxis
            .SetElement msoElementPrimaryValueAxisTitleAdjacentToAxis
            
            
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.caption = "挠度值（mm）"
             .AutoScaling = True     '自动坐标轴范围
    
            'ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "挠度值（mm）"
            
    End With
    

'    MS EXCEL VBA 代码2
'    Sheets(DispSheetName).Shapes.AddChart2(332, xlLineMarkersStacked, xPos, yPos + (i - 1) * yStep).Select
'    With ActiveChart
'
'            .SetSourceData Source:=Union(Range(Cells(curr, Node_Name_Col), Cells(curr + DispUbound(i) - 1, Node_Name_Col)), _
'            Range(Cells(curr, ElasticCol), Cells(curr + DispUbound(i) - 1, ElasticCol)), Range(Cells(curr, TheoryDisp_Col), Cells(curr + DispUbound(i) - 1, TheoryDisp_Col)))
'
'            .SetElement (msoElementChartTitleNone)    '删除标题
'            .SeriesCollection(1).Name = CStr(Cells(11, 9))
'            .SeriesCollection(2).Name = CStr(Cells(11, 10))
'
'            '横坐标
'            .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
'            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "测点号"
'            '纵坐标
'            .SetElement msoElementPrimaryValueAxisTitleBelowAxis
'            .Axes(xlValue, xlPrimary).AxisTitle.Text = "挠度值（mm）"
'
'    End With
        
        curr = curr + DispUbound(i)
    Next i
    Set plot = Nothing
    'ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    
        'ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    'ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
        'ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "挠度值（mm）"
    'Selection.Format.TextFrame2.TextRange.Characters.Text = "挠度值（mm）"
'    ActiveSheet.Shapes.AddChart2(332, xlLineMarkersStacked, 800, 150 + 220).Select
'    ActiveChart.SetSourceData Source:=Union(Range(Cells(13, 2), Cells(26, 2)), Range(Cells(13, 9), Cells(26, 9)), Range(Cells(13, 10), Cells(26, 10)))
'    ActiveChart.SetElement (msoElementChartTitleNone)
End Sub


