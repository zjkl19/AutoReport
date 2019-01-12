Attribute VB_Name = "AutoStrain"
Option Explicit

Private Const First_Row As Integer = 15    '��ʼ��������

Private Const MaxElasticStrain_Row As Integer = 4
Private Const MinStrainCheckoutCoff_Row As Integer = 5
Private Const MaxStrainCheckoutCoff_Row As Integer = 6
Private Const MinRefRemainStrain_Row As Integer = 7
Private Const MaxRefRemainStrain_Row As Integer = 8

Public StrainStatPara(1 To MAX_NWC, 1 To 5)  'ͳ�Ʋ���

Private Const e1 As Integer = 3  '��ʼģ���������У������ԣ�
Private Const e2 As Integer = 4  '��ʼ�¶�
Private Const e3 As Integer = 5  '����ģ��
Private Const e4 As Integer = 6  '�����¶�
Private Const e5 As Integer = 7  'ж��ģ��
Private Const e6 As Integer = 8  'ж���¶�

Const FullLoadStrainR_Col As Integer = 14
Const FullLoadStrainT_Col As Integer = 15
Const UnLoadStrainR_Col As Integer = 17
Const UnLoadStrainT_Col As Integer = 18

Dim c1 As Integer    '����ģ�����㣨�����У������ԣ�
Dim c2 As Integer    '�����¶ȼ���
Dim c3 As Integer    '����Ӧ��
Dim c4 As Integer    'ж��ģ������
Dim c5 As Integer    'ж���¶ȼ���
Dim c6 As Integer    'ж�ز���Ӧ��
Dim c7 As Integer    '����Ӧ��

Dim d1 As Integer    'ʵ����Ӧ�䣨�����У������ԣ�
Dim d2 As Integer    '����Ӧ��
Dim d3 As Integer    '����Ӧ��
Dim d4 As Integer    '��������ֵ
Dim d5 As Integer    'У��ϵ��
Dim d6 As Integer    '��Բ���Ӧ��


Const StrainNode_Name_Col As Integer = 2  '�����������

Public StrainGlobalWC(1 To MAX_NWC)    'ȫ�ֹ�����λ����

Public StrainNodeName(1 To MAX_NWC, 1 To MAX_NPS) As String  '���������������

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

Public TotalStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double    '��Ӧ��
Public RemainStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double    '����Ӧ��
Public ElasticStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double
Public TheoryStress(1 To MAX_NWC, 1 To MAX_NPS) As Double
Public TheoryStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double
Public StrainCheckoutCoff(1 To MAX_NWC, 1 To MAX_NPS) As Double
Public RefRemainStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double

Public INTTotalStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double '������Ӧ�䣨ȡ������ͬ��
Public INTElasticStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double
Public INTRemainStrain(1 To MAX_NWC, 1 To MAX_NPS) As Double
Public INTDivStrainCheckoutCoff(1 To MAX_NWC, 1 To MAX_NPS) As Double   '����Ӧ��ֵ�������ý������ͬ��
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

Public StrainUbound(1 To MAX_NWC) As Integer    'ÿ�������Ͻ磨�½�Ϊ1��
Public StrainNWCs As Integer    'Ӧ�乤����
Public StrainNPs(10) As Integer    '�������������
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
'����Ӧ�������
Private Sub GenerateStrainRows_Click()

    StrainNWCs = Cells(1, 2)
    
    Dim i As Integer
    For i = 0 To StrainNWCs - 1    'ÿ�����������
        StrainNPs(i) = Cells(2, 2 * (i + 1))
    Next
      
    Dim colorArray(1 To 9) As Integer
    For i = 1 To 8
        colorArray(i) = i
    Next
    colorArray(9) = TheoryStrain_Col
    
    Dim rowCurr As Integer    '��ָ��
    rowCurr = First_Row
    
    Dim ds As DeformService
    Set ds = New DeformService
    ds.GenerateRows StrainNWCs, StrainNPs, rowCurr, 1, colorArray
    Set ds = Nothing
 
End Sub

'�������
Public Sub StrainDataClear()
  If (MsgBox("����������ݲ��ɳ�������ȷ��Ҫ�����", vbYesNo + vbExclamation, "�ò������ɳ���") = vbNo) Then
    Exit Sub
  End If
  
  Dim i, j, k As Integer
  Dim rowCurr As Integer    '��ָ��
  rowCurr = First_Row
  
  'TODO:�����ʼ�������
  Dim dataArray(1 To 100) As Integer    '����յ���
  For i = 1 To 8
    dataArray(i) = i
  Next i
  For i = 12 To 21
    dataArray(i - 3) = i    '��Ŵ�9��ʼ
  Next i
  
  For i = 25 To 32
    dataArray(i - (25 - 18 - 1)) = i '��Ŵ�19��ʼ
  Next i
  
  
  '��ձ������
  While Cells(rowCurr, 1) <> ""    '��һ����Ԫ��������Ϊ�ж�����
    For i = 1 To UBound(dataArray)
        If dataArray(i) = 0 Then Exit For
        Cells(rowCurr, dataArray(i)) = ""
        Cells(rowCurr, dataArray(i)).Interior.Color = RGB(255, 255, 255) ' RGB(0, 176, 80)
    Next
    rowCurr = rowCurr + 1
  Wend

  '���ͳ������
    For i = 1 To MAX_NWC
        Cells(MaxElasticStrain_Row, 2 * i) = ""
        Cells(MinStrainCheckoutCoff_Row, 2 * i) = ""
        Cells(MaxStrainCheckoutCoff_Row, 2 * i) = ""
        Cells(MinRefRemainStrain_Row, 2 * i) = ""
        Cells(MaxRefRemainStrain_Row, 2 * i) = ""
    Next
End Sub

'����Ӧ��
'r2:�仯��ģ����r1:�仯ǰģ����t2:�仯���¶ȣ�t1���仯ǰ�¶�
Public Function GetStrain(ByVal r2, ByVal r1, ByVal t2, ByVal t1)
    Dim G, k, c
    G = 3.7: k = 1.8: c = 1.020019
    GetStrain = G * c * (r2 - r1) + k * (t2 - t1)
End Function

'�������Ӧ��
'deltaS:ж��״̬���ʼ״̬�ԱȲ�����Ӧ��ֵ��totalS:��Ӧ��
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

'�Զ�����Ӧ��
Public Sub AutoStrain_Click()

    InitStrainVar
    
    Dim i, j As Integer

    Dim rowCurr As Integer    '��ָ��
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
            Cells(rowCurr, INTRemainStrain_Col) = INTRemainStrain(i, j)    '����Ӧ��
            
            ElasticStrain(i, j) = TotalStrain(i, j) - RemainStrain(i, j)
            INTElasticStrain(i, j) = INTTotalStrain(i, j) - INTRemainStrain(i, j)
            Cells(rowCurr, INTElasticStrain_Col) = INTElasticStrain(i, j)    '����Ӧ��
             
            TheoryStrain(i, j) = Cells(rowCurr, TheoryStrain_Col)    '����Ӧ��ֱ��ȡֵ
            
            StrainCheckoutCoff(i, j) = ElasticStrain(i, j) / TheoryStrain(i, j)
            INTDivStrainCheckoutCoff(i, j) = INTElasticStrain(i, j) / TheoryStrain(i, j)
            Cells(rowCurr, StrainCheckoutCoff_Col) = INTDivStrainCheckoutCoff(i, j)    'У��ϵ��
            
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
            Cells(rowCurr, RefRemainStrain_Col) = INTDivRefRemainStrain(i, j)    '��Բ������
            
            rowCurr = rowCurr + 1
        Next
    Next
    
    'ע�⣡��Ӧ�䡢����Ӧ�䡢����Ӧ����ֵ��Ϊȡ�������ģ�
    '�������������С/��У��ϵ���������Բ���Ӧ��
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

        '����д��Excel
        Cells(MaxElasticStrain_Row, 2 * i) = Format(StrainStatPara(i, MaxElasticDeform_Index), "Fixed"): Cells(MinStrainCheckoutCoff_Row, 2 * i) = Format(StrainStatPara(i, MinCheckoutCoff_Index), "Fixed")
        Cells(MaxStrainCheckoutCoff_Row, 2 * i) = Format(StrainStatPara(i, MaxCheckoutCoff_Index), "Fixed"): Cells(MinRefRemainStrain_Row, 2 * i) = Format(StrainStatPara(i, MinRefRemainDeform_Index), "Percent")
        Cells(MaxRefRemainStrain_Row, 2 * i) = Format(StrainStatPara(i, MaxRefRemainDeform_Index), "Percent")
    Next
 
End Sub

'�Զ���ͼ����Ҫ�ȼ��㣩
Public Sub AutoStrainGraph()
    'https://docs.microsoft.com/zh-CN/office/vba/api/Excel.shapes.addchart2
    'AddChart2(��ʽ�� XlChartType�� �� ������ ��ȣ� �߶ȣ� NewLayout)
    Dim StrainSheetName As String
    StrainSheetName = "Ӧ��"
    
    Dim xPos, yPos As Integer  '�����һ��ͼx,yλ��
    xPos = 800: yPos = 150
    Dim yStep As Integer    'y����ÿ��ͼ��ռ�ÿռ�
    yStep = 220
      
   
    Dim i As Integer
    Dim curr As Integer
    curr = First_Row
    
    For i = 1 To StrainNWCs
        Sheets(StrainSheetName).Shapes.AddChart2(332, xlLineMarkersStacked, xPos, yPos + (i - 1) * yStep, 350, 200).Select
        With ActiveChart
    
                .SetSourceData Source:=Union(Range(Cells(curr, StrainNode_Name_Col), Cells(curr + StrainUbound(i) - 1, StrainNode_Name_Col)), _
                Range(Cells(curr, INTElasticStrain_Col), Cells(curr + StrainUbound(i) - 1, INTElasticStrain_Col)), Range(Cells(curr, TheoryStrain_Col), Cells(curr + StrainUbound(i) - 1, TheoryStrain_Col)))
                
                .SetElement (msoElementChartTitleNone)    'ɾ������
                .SeriesCollection(1).Name = "ʵ��ֵ"
                .SeriesCollection(2).Name = "����ֵ"
        
                '������
                .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
                .Axes(xlCategory, xlPrimary).AxisTitle.Text = "����"
                '������
                .SetElement msoElementPrimaryValueAxisTitleAdjacentToAxis
                
                .Axes(xlValue).HasTitle = True
                .Axes(xlValue).AxisTitle.Caption = "Ӧ�䣨�̦ţ�"
        
        End With
        
        curr = curr + DispUbound(i)
    Next i


End Sub
