VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Disp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_NWC As Integer = 10     '最大工况数
Private Const MAX_NPS As Integer = 100     '每个工况最大测点数

Private pDispUbound(1 To MAX_NWC) As Integer    '每个工况测点数上界（下界为1）
Private pGlobalWC(1 To MAX_NWC)  As Integer   '全局工况定位数组
Private pGroupName(1 To MAX_NWC)  As String '每个组的名称

Private pNodeName(1 To MAX_NWC, 1 To MAX_NPS) As String  '各个工况测点名称
Private pInitDisp(1 To MAX_NWC, 1 To MAX_NPS) As Double  '初始读数
Private pFullLoadDisp(1 To MAX_NWC, 1 To MAX_NPS)  '满载读数
Private pUnLoadDisp(1 To MAX_NWC, 1 To MAX_NPS)  '卸载读数
Private pTotalDisp(1 To MAX_NWC, 1 To 100)   'TotalDisp(i,j)表示第i个工况，第j个测点总变形
Private pDelta(1 To MAX_NWC, 1 To 100)    '增量，备用变量
Private pRemainDisp(1 To MAX_NWC, 1 To 100)
Private pElasticDisp(1 To MAX_NWC, 1 To 100)
Private pTheoryDisp(1 To MAX_NWC, 1 To 100)
Private pCheckoutCoff(1 To MAX_NWC, 1 To 100)
Private pRefRemainDisp(1 To MAX_NWC, 1 To 100)

Private pNWCs As Integer    '组数

'InitDisp(i, j) = Cells(rowCurr, InitDisp_Col)

Public Sub Init(ByVal NWCs As Integer, ByRef DispUbound() As Integer, ByRef GlobalWC() As Integer, ByRef GroupName() As String _
     , ByRef NodeName() As String)

    nPN = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十")
    pNWCs = NWCs
    Dim i As Integer: Dim j As Integer   '循环用变量
    For i = 1 To NWCs
        pDispUbound(i) = DispUbound(i)
    Next
    For i = 1 To NWCs
        pGlobalWC(i) = GlobalWC(i)
    Next
   For i = 1 To NWCs
        pGroupName(i) = GroupName(i)
    Next
    For i = 1 To NWCs
        For j = 1 To DispUbound(i)
            pNodeName(i, j) = NodeName(i, j)
        Next
    Next
'    For i = 1 To NWCs
'        GlobalWC(i) = Cells(3, 2 * i)
'    Next
'
'   For i = 1 To NWCs
'        GroupName(i) = Cells(4, 2 * i)
'    Next
End Sub

Public Property Get NWCs() As Integer
    NWCs = pNWCs
End Property

Public Property Get GlobalWC(ByVal i As Integer) As Integer
    GlobalWC = pGlobalWC(i)
End Property

Public Property Get GroupName(ByVal i As Integer) As String
    GroupName = pGroupName(i)
End Property

Public Property Get DispUbound(ByVal i As Integer) As Integer
    DispUbound = pDispUbound(i)
End Property

Public Property Get NodeName(ByVal i As Integer, ByVal j As Integer) As String
    NodeName = pNodeName(i, j)
End Property

Public Property Get InitDisp(ByVal i As Integer, ByVal j As Integer) As Double
    InitDisp = pInitDisp(i, j)
End Property

'TODO:只读属性
Public Property Let InitDisp(ByVal i As Integer, ByVal j As Integer, ByVal myVal As Double)
   pInitDisp(i, j) = myVal
End Property
