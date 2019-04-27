Attribute VB_Name = "AutoDisp"
Option Explicit
Const WC_Col As Integer = 1     '������������
Public Const MAX_NWC As Integer = 10     '��󹤿���
Public Const MAX_NPS As Integer = 100     'ÿ�������������

Private Const First_Row As Integer = 13     '��ʼ��������

Private Const MaxElasticDisp_Row As Integer = 5
Private Const MinCheckoutCoff_Row As Integer = 6
Private Const MaxCheckoutCoff_Row As Integer = 7
Private Const MinRefRemainDisp_Row As Integer = 8
Private Const MaxRefRemainDisp_Row As Integer = 9

Public StatPara(1 To MAX_NWC, 1 To 5)  'ͳ�Ʋ���


Const Node_Name_Col As Integer = 2  '����������У���ͬ��

Const InitDisp_Col As Integer = 3  '��ʼ����
Const FullLoadDisp_Col As Integer = 4  '���ض���
Const UnLoadDisp_Col As Integer = 5  'ж�ض���
Const TheoryDisp_Col As Integer = 6  '����λ��������

Const TotalDispCol As Integer = 16   '�ܱ���������
Const ElasticCol As Integer = 17   '���Ա���������
Const RemainDispCol As Integer = 18  '�������������
Private Const TheoryDispCopy_Col As Integer = 19  '��������
Const CheckoutCoffCol As Integer = 20   'У��ϵ��������
Const RefRemainDispCol As Integer = 21  '��Բ������������

Const DeltaCol As Integer = 22  '����������

Const GroupNameCol As Integer = 13   '��������������

Public NWCs As Integer    '���Ӷȣ�������

Public GlobalWC(1 To MAX_NWC)  As Integer   'ȫ�ֹ�����λ����
Public GroupName(1 To MAX_NWC)  As String 'ÿ���������

'Obsolete code:
'Public GroupName(1 To MAX_NWC, 1 To MAX_NPS)  As String '�����������������������
'��GroupName(i,j)��ʾ��i��������j�����������������

Public InitDisp(1 To MAX_NWC, 1 To MAX_NPS)  '��ʼ����
Public FullLoadDisp(1 To MAX_NWC, 1 To MAX_NPS)  '���ض���
Public UnLoadDisp(1 To MAX_NWC, 1 To MAX_NPS)  'ж�ض���

Public TotalDisp(1 To MAX_NWC, 1 To 100)   'TotalDisp(i,j)��ʾ��i����������j������ܱ���
Public NodeName(1 To MAX_NWC, 1 To 100) As String  '���������������
Public Delta(1 To MAX_NWC, 1 To 100)    '���������ñ���
Public RemainDisp(1 To MAX_NWC, 1 To 100)
Public ElasticDisp(1 To MAX_NWC, 1 To 100)
Public TheoryDisp(1 To MAX_NWC, 1 To 100)
Public CheckoutCoff(1 To MAX_NWC, 1 To 100)
Public RefRemainDisp(1 To MAX_NWC, 1 To 100)
Public DispUbound(1 To MAX_NWC) As Integer    'ÿ�������Ͻ磨�½�Ϊ1��

Public DispChartObjArray(1 To MAX_NWC) As Shape    'as Object    '�洢��ͼ��ָ��

Public dispResultVar(1 To MAX_NWC) As String    '��Ӧ�ĵ�������
Public dispResult(1 To MAX_NWC) As String    '��word����ӦDocVariable��Ӧ
Public dispSummaryVar(1 To MAX_NWC) As String    '��Ӧ�ĵ�������
Public dispSummary(1 To MAX_NWC) As String    '��word����ӦDocVariable��Ӧ
Public dispTbTitle(1 To MAX_NWC) As String
Public dispTbCrossRef(1 To MAX_NWC) As String    '���潻������
Public dispGraphCrossRef(1 To MAX_NWC) As String

Public dispRawTbTitle(1 To MAX_NWC) As String
Public dispChartTitle(1 To MAX_NWC) As String

Public dispTheoryShapeTitle(1 To MAX_NWC) As String    '�����Ӷ�ֵ
Public dispTheoryShapeTitleVar(1 To MAX_NWC) As String

Public dispRawTblBookmarks(1 To MAX_NWC) As String
Public dispTblBookmarks(1 To MAX_NWC) As String
Public dispChartBookmarks(1 To MAX_NWC) As String
Public dispTheoryShapeBookmarks(1 To MAX_NWC) As String


'�������
Public Sub DispDataClear()
  If (MsgBox("����������ݲ��ɳ�������ȷ��Ҫ�����", vbYesNo + vbExclamation, "�ò������ɳ���") = vbNo) Then
    Exit Sub
  End If
  
  Dim i, j As Integer
  Dim rowCurr As Integer    '��ָ��
  rowCurr = First_Row
  
  '��ձ������
  While Cells(rowCurr, 1) <> ""    '��һ����Ԫ��������Ϊ�ж�����
    For i = 1 To TheoryDisp_Col
        Cells(rowCurr, i) = ""
        Cells(rowCurr, i).Interior.Color = RGB(255, 255, 255) ' RGB(0, 176, 80)
    Next
    rowCurr = rowCurr + 1
  Wend
  
  '���ͳ������
    For i = 1 To MAX_NWC
        Cells(MaxElasticDisp_Row, 2 * i) = ""
        Cells(MinCheckoutCoff_Row, 2 * i) = ""
        Cells(MaxCheckoutCoff_Row, 2 * i) = ""
        Cells(MinRefRemainDisp_Row, 2 * i) = ""
        Cells(MaxRefRemainDisp_Row, 2 * i) = ""
    Next
End Sub
'''��ʼ��ȫ�ֱ���
Public Sub InitVar()

    nPN = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ")
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
    
    Dim ax As ArrayService    'ԭas��������ؼ���as��ͻ����Ϊas
    Set ax = New ArrayService
    
    Dim rowCurr As Integer    '��ָ��
    Dim i, j As Integer
    rowCurr = First_Row
    For i = 1 To NWCs
        For j = 1 To DispUbound(i)
        
            NodeName(i, j) = Cells(rowCurr, Node_Name_Col)
            
            TheoryDisp(i, j) = Cells(rowCurr, TheoryDisp_Col)
            
            TheoryDisp(i, j) = Round(TheoryDisp(i, j), 2)    'ȡ2λС��
            Cells(rowCurr, TheoryDispCopy_Col) = TheoryDisp(i, j)
            
            InitDisp(i, j) = Cells(rowCurr, InitDisp_Col)
            FullLoadDisp(i, j) = Cells(rowCurr, FullLoadDisp_Col)
            UnLoadDisp(i, j) = Cells(rowCurr, UnLoadDisp_Col)
            
            
            TotalDisp(i, j) = FullLoadDisp(i, j) - InitDisp(i, j) '�ܱ���
            Cells(rowCurr, TotalDispCol) = TotalDisp(i, j)
            
            Cells(rowCurr, DeltaCol) = UnLoadDisp(i, j) - FullLoadDisp(i, j) '����
            '�������洢
            
            '�㷨��ж�������ض�����ֵ>=0��ȡж�������ض�����ֵ������ȡ0
            'RemainDisp(i, j) = IIf(UnLoadDisp(i, j) - InitDisp(i, j) >= 0, UnLoadDisp(i, j) - InitDisp(i, j), 0)
            RemainDisp(i, j) = ds.GetRemainDisp(InitDisp(i, j), FullLoadDisp(i, j), UnLoadDisp(i, j))
            Cells(rowCurr, RemainDispCol) = RemainDisp(i, j)    '�������
            
            ElasticDisp(i, j) = TotalDisp(i, j) - RemainDisp(i, j)
            Cells(rowCurr, ElasticCol) = ElasticDisp(i, j)    '���Ա���
             
            CheckoutCoff(i, j) = ElasticDisp(i, j) / TheoryDisp(i, j)
            Cells(rowCurr, CheckoutCoffCol) = CheckoutCoff(i, j)    'У��ϵ��
             
            RefRemainDisp(i, j) = RemainDisp(i, j) / TotalDisp(i, j)
            Cells(rowCurr, RefRemainDispCol) = RefRemainDisp(i, j)    '��Բ������
            
            rowCurr = rowCurr + 1
        Next
    Next
    
    '�����������ͳ�Ʋ���
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
        
        '����д��Excel
        Cells(MaxElasticDisp_Row, 2 * i) = Format(StatPara(i, 1), "Fixed"): Cells(MinCheckoutCoff_Row, 2 * i) = Format(StatPara(i, 2), "Fixed")
        Cells(MaxCheckoutCoff_Row, 2 * i) = Format(StatPara(i, MaxCheckoutCoff_Index), "Fixed"): Cells(MinRefRemainDisp_Row, 2 * i) = Format(StatPara(i, MinRefRemainDeform_Index), "Percent")
        Cells(MaxRefRemainDisp_Row, 2 * i) = Format(StatPara(i, MaxRefRemainDeform_Index), "Percent")
    Next
    
    '�����������Ӧ�ĵ�������ֵ
    '�㷨�����ݱ��湤������������
    For i = 1 To NWCs
         '�����Ӧ�ĵ�����
         dispResultVar(i) = Replace("dispResult" & CStr(i), " ", "")
         dispSummaryVar(i) = Replace("dispSummary" & CStr(i), " ", "")
         dispTbCrossRef(i) = "dispTbCrossRef" & CStr(i)
         dispGraphCrossRef(i) = "dispGraphCrossRef" & CStr(i)
         
         '��Ӧ��ǩ��
        dispRawTblBookmarks(i) = Replace("dispRawTb" & CStr(i), " ", "")
        dispTblBookmarks(i) = Replace("dispTable" & CStr(i), " ", "")
        dispChartBookmarks(i) = Replace("dispChart" & CStr(i), " ", "")
        dispTheoryShapeBookmarks(i) = Replace("dispTheoryShape" & CStr(i), " ", "")
        
        '��Ӧ�ĵ�����
        dispTheoryShapeTitleVar(i) = Replace("dispTheoryShapeTitle" & CStr(i), " ", "")
        
        '�����������Լ�ͼ�����
        If ax.CountElements(GlobalWC, GlobalWC(i)) = 1 Then    '����ֻ��1��
            dispResult(i) = "(" & CStr(i) & ")�ڹ���" & CStr(nPN(GlobalWC(i) - 1)) & "���������£��������ʵ�ⵯ���Ӷ�ֵΪ" & Format(StatPara(i, MaxElasticDeform_Index), "Fixed") & "mm��" _
                & "ʵ����ƽ�����Ӷ�ֵ��С������ֵ��У��ϵ����" & Format(StatPara(i, MinCheckoutCoff_Index), "Fixed") & "��" & Format(StatPara(i, MaxCheckoutCoff_Index), "Fixed") & "֮�䣻" _
                & "��Բ��������" & Format(StatPara(i, MinRefRemainDeform_Index), "Percent") & "��" & Format(StatPara(i, MaxRefRemainDeform_Index), "Percent") & "֮�䡣"
            dispSummary(i) = "����" & CStr(nPN(GlobalWC(i) - 1)) & "�����Ӷȼ�������" & dispTbCrossRef(i) & "���Ӷ�ʵ��ֵ�����ۼ���ֵ�Ĺ�ϵ�������" & dispGraphCrossRef(i) & "������������������������Ӷ�У��ϵ����" & Format(StatPara(i, MinCheckoutCoff_Index), "Fixed") & "��" & Format(StatPara(i, MaxCheckoutCoff_Index), "Fixed") & "֮�䣬" _
                & "��Բ��������" & Format(StatPara(i, MinRefRemainDeform_Index), "Percent") & "��" & Format(StatPara(i, MaxRefRemainDeform_Index), "Percent") & "֮�䡣"
            dispTbTitle(i) = "����" & CStr(nPN(GlobalWC(i) - 1)) & "�Ӷȼ�������ܱ�"
            dispChartTitle(i) = "����" & CStr(nPN(GlobalWC(i) - 1)) & "�Ӷ�ʵ��ֵ�����ۼ���ֵ�Ĺ�ϵ����"
            
            dispTheoryShapeTitle(i) = "����" & CStr(nPN(GlobalWC(i) - 1)) & "���ƽ��������Ӷ�ֵ����λ��mm��"
            dispRawTbTitle(i) = "����" & CStr(nPN(GlobalWC(i) - 1)) & "�Ӷ�ԭʼ���ݴ����"    '���������ʹ�õı���
        Else
             dispResult(i) = "(" & CStr(i) & ")�ڹ���" & CStr(nPN(GlobalWC(i) - 1)) & "���������£�����" & GroupName(i) & "�������ʵ�ⵯ���Ӷ�ֵΪ" & Format(StatPara(i, MaxElasticDeform_Index), "Fixed") & "mm��" _
                & "ʵ����ƽ�����Ӷ�ֵ��С������ֵ��У��ϵ����" & Format(StatPara(i, MinCheckoutCoff_Index), "Fixed") & "��" & Format(StatPara(i, MaxCheckoutCoff_Index), "Fixed") & "֮�䣻" _
                & "��Բ��������" & Format(StatPara(i, MinRefRemainDeform_Index), "Percent") & "��" & Format(StatPara(i, MaxRefRemainDeform_Index), "Percent") & "֮�䡣"
            dispSummary(i) = "����" & CStr(nPN(GlobalWC(i) - 1)) & "����" & GroupName(i) & "�����Ӷȼ�������" & dispTbCrossRef(i) & "���Ӷ�ʵ��ֵ�����ۼ���ֵ�Ĺ�ϵ�������" & dispGraphCrossRef(i) & "������������������������Ӷ�У��ϵ����" & Format(StatPara(i, MinCheckoutCoff_Index), "Fixed") & "��" & Format(StatPara(i, MaxCheckoutCoff_Index), "Fixed") & "֮�䣬" _
                & "��Բ��������" & Format(StatPara(i, MinRefRemainDeform_Index), "Percent") & "��" & Format(StatPara(i, MaxRefRemainDeform_Index), "Percent") & "֮�䡣"
            dispTbTitle(i) = "����" & CStr(nPN(GlobalWC(i) - 1)) & GroupName(i) & "�����Ӷȼ�������ܱ�"
            dispChartTitle(i) = "����" & CStr(nPN(GlobalWC(i) - 1)) & GroupName(i) & "�����Ӷ�ʵ��ֵ�����ۼ���ֵ�Ĺ�ϵ����"
            
            dispTheoryShapeTitle(i) = "����" & CStr(nPN(GlobalWC(i) - 1)) & GroupName(i) & "���������Ӷ�ֵ����λ��mm��"
            dispRawTbTitle(i) = "����" & CStr(nPN(GlobalWC(i) - 1)) & GroupName(i) & "�����Ӷ�ԭʼ���ݴ����"    '���������ʹ�õı���
        End If

    Next
    Set ds = Nothing
    Set gs = Nothing
    Set ax = Nothing
End Sub

Private Sub GenerateRows_Click()
    Dim NWCs As Integer    '������
    Dim nPs(10) As Integer    '�������������
    Dim nPN     '����������Ӧ��������
    nPN = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ")
    Dim i, j As Integer
    NWCs = Cells(1, 2)
    For i = 0 To NWCs - 1
        nPs(i) = Cells(2, 2 * (i + 1))
        'Debug.Print nPs(i)
    Next
    'Debug.Print nWCs
       
    Dim rowCurr As Integer    '��ָ��
    rowCurr = First_Row
    
    For i = 0 To NWCs - 1    '��������
        For j = 1 To nPs(i)    '�������������Ĳ��
            Cells(rowCurr, WC_Col) = nPN(i)
            
            '���ñ�����ı���ɫ
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
    
    'TODO:��ӱ���ɫ
    
End Sub

'�Զ���ͼ����Ҫ�ȼ��㣩
Private Sub AutoGraph_Click()
    'https://docs.microsoft.com/zh-CN/office/vba/api/Excel.shapes.addchart2
    'AddChart2(��ʽ�� XlChartType�� �� ������ ��ȣ� �߶ȣ� NewLayout)
    Dim DispSheetName As String
    DispSheetName = "�Ӷ�"
    
    Dim i As Integer
    Dim xPos, yPos As Integer  '�����һ��ͼx,yλ��
    xPos = 400: yPos = 150
    Dim yStep As Integer    'y����ÿ��ͼ��ռ�ÿռ�
    yStep = 260    'ԭΪ220������Ϊ���޼���
    
    Dim chartWidth As Integer: Dim chartHeight As Integer
   
    Dim curr As Integer: Dim currCounts As Integer   '��ǰ��ͼ�����
    curr = First_Row: currCounts = 1 '����һ����
    
    Dim plot As Excel.Shape

    'Set plot = ws.Shapes.AddChart
    
    For i = 1 To NWCs
'        MS EXCEL VBA ����1
'        Set plot = Sheets(DispSheetName).Shapes.AddChart2(332, xlLineMarkersStacked, xPos, yPos + (i - 1) * yStep)
'        plot.Chart.SetSourceData Source:=Union(Range(Cells(curr, Node_Name_Col), Cells(curr + DispUbound(i) - 1, Node_Name_Col)), _
'        Range(Cells(curr, ElasticCol), Cells(curr + DispUbound(i) - 1, ElasticCol)), Range(Cells(curr, TheoryDisp_Col), Cells(curr + DispUbound(i) - 1, TheoryDisp_Col)))
'
'        plot.Chart.SetElement (msoElementChartTitleNone)    'ɾ������
'
'        plot.Chart.SeriesCollection(1).Name = CStr(Cells(11, 9))
'        plot.Chart.SeriesCollection(2).Name = CStr(Cells(11, 10))
'
'        '������
'        plot.Chart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
'        plot.Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "����"
'        '������
'        plot.Chart.SetElement msoElementPrimaryValueAxisTitleBelowAxis
'        plot.Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "�Ӷ�ֵ��mm��"

    currCounts = Range(Cells(curr, Node_Name_Col), Cells(curr + DispUbound(i) - 1, Node_Name_Col)).Rows.Count
    chartWidth = 350: chartHeight = 220 'Ĭ��ֵ
    If currCounts > 12 And currCounts < 15 Then
        chartWidth = 380: chartHeight = 240
    ElseIf currCounts >= 15 And currCounts < 23 Then
         chartWidth = 410: chartHeight = 250
    ElseIf currCounts >= 23 Then
         chartWidth = 450: chartHeight = 270    '�߿��������������ͬ
    End If
    Set DispChartObjArray(i) = Sheets(DispSheetName).Shapes.AddChart2(332, xlLineMarkers, xPos, yPos + (i - 1) * yStep, chartWidth, chartHeight)
    DispChartObjArray(i).Select
    'Դ���룺'Sheets(DispSheetName).Shapes.AddChart2(332, xlLineMarkers, xPos, yPos + (i - 1) * yStep, 350, 200).Select
    With ActiveChart
            
            .SetSourceData Source:=Union(Range(Cells(curr, Node_Name_Col), Cells(curr + DispUbound(i) - 1, Node_Name_Col)), _
            Range(Cells(curr, ElasticCol), Cells(curr + DispUbound(i) - 1, ElasticCol)), Range(Cells(curr, TheoryDisp_Col), Cells(curr + DispUbound(i) - 1, TheoryDisp_Col)))
            
            
            .SetElement (msoElementChartTitleNone)    'ɾ������
            .SeriesCollection(1).name = "����ֵ" 'CStr(Cells(11, 9))
            .SeriesCollection(2).name = "ʵ��ֵ" ' CStr(Cells(11, 10))
            .Axes(xlValue).TickLabels.NumberFormatLocal = "#,##0.00_ "

            '������
             '.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
             '.Axes(xlCategory, xlPrimary).AxisTitle.Text = "����"
                
            '������
            'ActiveChart.SetElement msoElementPrimaryValueAxisTitleBelowAxis
            .SetElement msoElementPrimaryValueAxisTitleAdjacentToAxis
            
            
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.caption = "�Ӷ�ֵ��mm��"
             .AutoScaling = True     '�Զ������᷶Χ
    
            'ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "�Ӷ�ֵ��mm��"
            
    End With
    

'    MS EXCEL VBA ����2
'    Sheets(DispSheetName).Shapes.AddChart2(332, xlLineMarkersStacked, xPos, yPos + (i - 1) * yStep).Select
'    With ActiveChart
'
'            .SetSourceData Source:=Union(Range(Cells(curr, Node_Name_Col), Cells(curr + DispUbound(i) - 1, Node_Name_Col)), _
'            Range(Cells(curr, ElasticCol), Cells(curr + DispUbound(i) - 1, ElasticCol)), Range(Cells(curr, TheoryDisp_Col), Cells(curr + DispUbound(i) - 1, TheoryDisp_Col)))
'
'            .SetElement (msoElementChartTitleNone)    'ɾ������
'            .SeriesCollection(1).Name = CStr(Cells(11, 9))
'            .SeriesCollection(2).Name = CStr(Cells(11, 10))
'
'            '������
'            .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
'            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "����"
'            '������
'            .SetElement msoElementPrimaryValueAxisTitleBelowAxis
'            .Axes(xlValue, xlPrimary).AxisTitle.Text = "�Ӷ�ֵ��mm��"
'
'    End With
        
        curr = curr + DispUbound(i)
    Next i
    Set plot = Nothing
    'ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    
        'ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    'ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
        'ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "�Ӷ�ֵ��mm��"
    'Selection.Format.TextFrame2.TextRange.Characters.Text = "�Ӷ�ֵ��mm��"
'    ActiveSheet.Shapes.AddChart2(332, xlLineMarkersStacked, 800, 150 + 220).Select
'    ActiveChart.SetSourceData Source:=Union(Range(Cells(13, 2), Cells(26, 2)), Range(Cells(13, 9), Cells(26, 9)), Range(Cells(13, 10), Cells(26, 10)))
'    ActiveChart.SetElement (msoElementChartTitleNone)
End Sub


