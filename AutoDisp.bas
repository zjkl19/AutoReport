Attribute VB_Name = "AutoDisp"
Option Explicit
Const WC_Col As Integer = 1     '工况所在列数
Public Const MAX_NWC As Integer = 10     '最大工况数
Public Const MAX_NPS As Integer = 100     '每个工况最大测点数

Private Const First_Row As Integer = 13     '起始数据行数

Private Const MaxElasticDisp_Row As Integer = 4
Private Const MinCheckoutCoff_Row As Integer = 5
Private Const MaxCheckoutCoff_Row As Integer = 6
Private Const MinRefRemainDisp_Row As Integer = 7
Private Const MaxRefRemainDisp_Row As Integer = 8

Public StatPara(1 To MAX_NWC, 1 To 5)  '统计参数


Const Node_Name_Col As Integer = 2  '测点编号所在列（下同）

Const InitDisp_Col As Integer = 3  '初始读数
Const FullLoadDisp_Col As Integer = 4  '满载读数
Const UnLoadDisp_Col As Integer = 6  '卸载读数


Const TheoryDisp_Col As Integer = 10  '理论位移所在列
Const TotalDispCol As Integer = 5   '总变形所在列
Const DeltaCol As Integer = 7  '增量所在列
Const RemainDispCol As Integer = 8   '残余变形所在列
Const ElasticCol As Integer = 9   '弹性变形所在列
Const CheckoutCoffCol As Integer = 11   '校验系数所在列
Const RefRemainDispCol As Integer = 12   '相对残余变形所在列

Public nWCs As Integer    '（挠度）工况数
Public nPN    '各个工况对应中文名称
'nPN = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十")

Public GlobalWC(1 To MAX_NWC)    '全局工况定位数组

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
    For i = 1 To RefRemainDispCol
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
    nWCs = Cells(1, 2)
    Dim i As Integer
    For i = 1 To nWCs
        DispUbound(i) = Cells(2, 2 * i)
    Next
    

    For i = 1 To nWCs
        GlobalWC(i) = Cells(3, 2 * i)
    Next
    
End Sub


Public Sub AutoDisp_Click()

    InitVar
    
    Dim rowCurr As Integer    '行指针
    
    Dim i, j As Integer
    rowCurr = First_Row
    For i = 1 To nWCs
        For j = 1 To DispUbound(i)
        
            NodeName(i, j) = Cells(rowCurr, Node_Name_Col)
            TheoryDisp(i, j) = Cells(rowCurr, TheoryDisp_Col)
            
            TheoryDisp(i, j) = Round(TheoryDisp(i, j), 2)    '取2位小数
            
            InitDisp(i, j) = Cells(rowCurr, InitDisp_Col)
            FullLoadDisp(i, j) = Cells(rowCurr, FullLoadDisp_Col)
            UnLoadDisp(i, j) = Cells(rowCurr, UnLoadDisp_Col)
            
            
            TotalDisp(i, j) = Cells(rowCurr, TotalDispCol - 1) - Cells(rowCurr, TotalDispCol - 2)    '总变形
            Cells(rowCurr, TotalDispCol) = TotalDisp(i, j)
            
            Cells(rowCurr, DeltaCol) = Cells(rowCurr, DeltaCol - 1) - Cells(rowCurr, DeltaCol - 3)   '增量
            '增量不存储
            
            '算法：卸载与满载读数差值>=0，取卸载与满载读数差值，否则取0
            RemainDisp(i, j) = IIf(UnLoadDisp(i, j) - InitDisp(i, j) >= 0, UnLoadDisp(i, j) - InitDisp(i, j), 0)
            Cells(rowCurr, RemainDispCol) = RemainDisp(i, j)    '残余变形
            
            ElasticDisp(i, j) = Cells(rowCurr, ElasticCol - 4) - Cells(rowCurr, ElasticCol - 1)
            Cells(rowCurr, ElasticCol) = ElasticDisp(i, j)    '弹性变形
             
            CheckoutCoff(i, j) = Cells(rowCurr, CheckoutCoffCol - 2) / Cells(rowCurr, CheckoutCoffCol - 1)
            Cells(rowCurr, CheckoutCoffCol) = CheckoutCoff(i, j)    '校验系数
             
            RefRemainDisp(i, j) = Cells(rowCurr, RefRemainDispCol - 4) / Cells(rowCurr, RefRemainDispCol - 7)
            Cells(rowCurr, RefRemainDispCol) = RefRemainDisp(i, j)    '相对残余变形
            
            rowCurr = rowCurr + 1
        Next
    Next

    '计算各个工况统计参数
    For i = 1 To nWCs
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

 
End Sub

Private Sub GenerateRows_Click()
    Dim nWCs As Integer    '工况数
    Dim nPs(10) As Integer    '各个工况测点数
    Dim nPN     '各个工况对应中文名称
    nPN = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十")
    Dim i, j As Integer
    nWCs = Cells(1, 2)
    For i = 0 To nWCs - 1
        nPs(i) = Cells(2, 2 * (i + 1))
        'Debug.Print nPs(i)
    Next
    'Debug.Print nWCs
       
    Dim rowCurr As Integer    '行指针
    rowCurr = First_Row
    
    For i = 0 To nWCs - 1    '遍历工况
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
    
    Dim xPos, yPos As Integer  '定义第一张图x,y位置
    xPos = 800: yPos = 150
    Dim yStep As Integer    'y方向每个图表占用空间
    yStep = 220
    
    Dim plot As Excel.Shape
    

    'Set plot = ws.Shapes.AddChart
    
    Dim i As Integer
    Dim curr As Integer
    curr = First_Row
    
    For i = 1 To nWCs
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
    Sheets(DispSheetName).Shapes.AddChart2(332, xlLineMarkers, xPos, yPos + (i - 1) * yStep, 350, 200).Select
    With ActiveChart
            
            .SetSourceData Source:=Union(Range(Cells(curr, Node_Name_Col), Cells(curr + DispUbound(i) - 1, Node_Name_Col)), _
            Range(Cells(curr, ElasticCol), Cells(curr + DispUbound(i) - 1, ElasticCol)), Range(Cells(curr, TheoryDisp_Col), Cells(curr + DispUbound(i) - 1, TheoryDisp_Col)))
            
            
            .SetElement (msoElementChartTitleNone)    '删除标题
            .SeriesCollection(1).Name = "实测值" 'CStr(Cells(11, 9))
            .SeriesCollection(2).Name = "理论值" ' CStr(Cells(11, 10))
    
            '横坐标
            .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "测点号"
            '纵坐标
            'ActiveChart.SetElement msoElementPrimaryValueAxisTitleBelowAxis
            .SetElement msoElementPrimaryValueAxisTitleAdjacentToAxis
            
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Caption = "挠度值（mm）"
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


