Attribute VB_Name = "AutoDisp"
Option Explicit
Const WC_Col As Integer = 1     '������������
Public Const MAX_NWC As Integer = 10     '��󹤿���
Public Const MAX_NPS As Integer = 100     'ÿ�������������

Private Const First_Row As Integer = 13     '��ʼ��������

Private Const MaxElasticDisp_Row As Integer = 4
Private Const MinCheckoutCoff_Row As Integer = 5
Private Const MaxCheckoutCoff_Row As Integer = 6
Private Const MinRefRemainDisp_Row As Integer = 7
Private Const MaxRefRemainDisp_Row As Integer = 8

Public StatPara(1 To MAX_NWC, 1 To 5)  'ͳ�Ʋ���


Const Node_Name_Col As Integer = 2  '�����������
Const TheoryDisp_Col As Integer = 10  '����λ��������
Dim TotalDispCol As Integer    '�ܱ���������
Dim DeltaCol As Integer   '����������
Dim RemainDispCol As Integer    '�������������
Dim ElasticCol As Integer    '���Ա���������
Dim CheckoutCoffCol As Integer    'У��ϵ��������
Dim RefRemainDispCol As Integer    '��Բ������������

Public nWCs As Integer    '���Ӷȣ�������
Public nPN    '����������Ӧ��������
'nPN = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ")

Public GlobalWC(1 To MAX_NWC)    'ȫ�ֹ�����λ����

Public TotalDisp(1 To MAX_NWC, 1 To 100)   'TotalDisp(i,j)��ʾ��i����������j������ܱ���
Public NodeName(1 To MAX_NWC, 1 To 100) As String  '���������������
Public Delta(1 To MAX_NWC, 1 To 100)    '���������ñ���
Public RemainDisp(1 To MAX_NWC, 1 To 100)
Public ElasticDisp(1 To MAX_NWC, 1 To 100)
Public TheoryDisp(1 To MAX_NWC, 1 To 100)
Public CheckoutCoff(1 To MAX_NWC, 1 To 100)
Public RefRemainDisp(1 To MAX_NWC, 1 To 100)
Public DispUbound(1 To MAX_NWC) As Integer    'ÿ�������Ͻ磨�½�Ϊ1��



Dim t

'''��ʼ��ȫ�ֱ���
Public Sub InitVar()

    nPN = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ")
    nWCs = Cells(1, 2)
    Dim i As Integer
    For i = 1 To nWCs
        DispUbound(i) = Cells(2, 2 * i)
    Next
    

    For i = 1 To nWCs
        GlobalWC(i) = Cells(3, 2 * i)
    Next
    
    TotalDispCol = 5
    DeltaCol = 7
    RemainDispCol = 8
    ElasticCol = 9
    CheckoutCoffCol = 11
    RefRemainDispCol = 12
End Sub


Public Sub AutoDisp_Click()

    InitVar
    
    Dim rowCurr As Integer    '��ָ��
    
    Dim i, j As Integer
    rowCurr = First_Row
    For i = 1 To nWCs
        For j = 1 To DispUbound(i)
        
            NodeName(i, j) = Cells(rowCurr, Node_Name_Col)
            TheoryDisp(i, j) = Cells(rowCurr, TheoryDisp_Col)
            
            TotalDisp(i, j) = Cells(rowCurr, TotalDispCol - 1) - Cells(rowCurr, TotalDispCol - 2)    '�ܱ���
            Cells(rowCurr, TotalDispCol) = TotalDisp(i, j)
            
            Cells(rowCurr, DeltaCol) = Cells(rowCurr, DeltaCol - 1) - Cells(rowCurr, DeltaCol - 3)   '����
            '�������洢
            
            '�㷨��ж�������ض�����ֵ>=0��ȡж�������ض�����ֵ������ȡ0
            RemainDisp(i, j) = IIf(Cells(rowCurr, RemainDispCol - 2) - Cells(rowCurr, RemainDispCol - 5) >= 0, Cells(rowCurr, RemainDispCol - 2) - Cells(rowCurr, RemainDispCol - 5), 0)
            Cells(rowCurr, RemainDispCol) = RemainDisp(i, j)    '�������
            
            ElasticDisp(i, j) = Cells(rowCurr, ElasticCol - 4) - Cells(rowCurr, ElasticCol - 1)
            Cells(rowCurr, ElasticCol) = ElasticDisp(i, j)    '���Ա���
             
            CheckoutCoff(i, j) = Cells(rowCurr, CheckoutCoffCol - 2) / Cells(rowCurr, CheckoutCoffCol - 1)
            Cells(rowCurr, CheckoutCoffCol) = CheckoutCoff(i, j)    'У��ϵ��
             
            RefRemainDisp(i, j) = Cells(rowCurr, RefRemainDispCol - 4) / Cells(rowCurr, RefRemainDispCol - 7)
            Cells(rowCurr, RefRemainDispCol) = RefRemainDisp(i, j)    '��Բ������
            
            rowCurr = rowCurr + 1
        Next
    Next

    '�����������ͳ�Ʋ���
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
        
        '����д��Excel
        Cells(MaxElasticDisp_Row, 2 * i) = Format(StatPara(i, 1), "Fixed"): Cells(MinCheckoutCoff_Row, 2 * i) = Format(StatPara(i, 2), "Fixed")
        Cells(MaxCheckoutCoff_Row, 2 * i) = Format(StatPara(i, MaxCheckoutCoff_Index), "Fixed"): Cells(MinRefRemainDisp_Row, 2 * i) = Format(StatPara(i, MinRefRemainDeform_Index), "Percent")
        Cells(MaxRefRemainDisp_Row, 2 * i) = Format(StatPara(i, MaxRefRemainDeform_Index), "Percent")
    Next

 
End Sub

Private Sub GenerateRows_Click()
    Dim nWCs As Integer    '������
    Dim nPs(10) As Integer    '�������������
    Dim nPN     '����������Ӧ��������
    nPN = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ")
    Dim i, j As Integer
    nWCs = Cells(1, 2)
    For i = 0 To nWCs - 1
        nPs(i) = Cells(2, 2 * (i + 1))
        'Debug.Print nPs(i)
    Next
    'Debug.Print nWCs
       
    Dim rowCurr As Integer    '��ָ��
    rowCurr = First_Row
    
    For i = 0 To nWCs - 1    '��������
        For j = 1 To nPs(i)    '�������������Ĳ��
            Cells(rowCurr, WC_Col) = nPN(i)
            rowCurr = rowCurr + 1
        Next
    Next
End Sub

'�Զ���ͼ����Ҫ�ȼ��㣩
Private Sub AutoGraph_Click()
    'https://docs.microsoft.com/zh-CN/office/vba/api/Excel.shapes.addchart2
    'AddChart2(��ʽ�� XlChartType�� �� ������ ��ȣ� �߶ȣ� NewLayout)
    Dim xPos, yPos As Integer  '�����һ��ͼx,yλ��
    xPos = 800: yPos = 150
    Dim yStep As Integer    'y����ÿ��ͼ��ռ�ÿռ�
    yStep = 220
    
    Dim plot As Excel.Shape
    

    'Set plot = ws.Shapes.AddChart
    
    Dim i As Integer
    Dim curr As Integer
    curr = First_Row
    
    For i = 1 To nWCs

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
    Sheets(DispSheetName).Shapes.AddChart2(332, xlLineMarkersStacked, xPos, yPos + (i - 1) * yStep).Select
    With ActiveChart
            
            .SetSourceData Source:=Union(Range(Cells(curr, Node_Name_Col), Cells(curr + DispUbound(i) - 1, Node_Name_Col)), _
            Range(Cells(curr, ElasticCol), Cells(curr + DispUbound(i) - 1, ElasticCol)), Range(Cells(curr, TheoryDisp_Col), Cells(curr + DispUbound(i) - 1, TheoryDisp_Col)))
            
            .SetElement (msoElementChartTitleNone)    'ɾ������
            
            .SeriesCollection(1).Name = CStr(Cells(11, 9))
            .SeriesCollection(2).Name = CStr(Cells(11, 10))
    
            '������
            .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
            .Axes(xlCategory, xlPrimary).AxisTitle.Text = "����"
            '������
            .SetElement msoElementPrimaryValueAxisTitleBelowAxis
            .Axes(xlValue, xlPrimary).AxisTitle.Text = "�Ӷ�ֵ��mm��"
            
    End With
        
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


