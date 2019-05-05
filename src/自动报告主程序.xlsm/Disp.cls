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

Private pInitDisp(1 To MAX_NWC, 1 To MAX_NPS) As Double  '初始读数

'InitDisp(i, j) = Cells(rowCurr, InitDisp_Col)

Public Sub Init(ByVal sheetName As String)
   pInitDisp(1, 1) = CDbl(Sheets(sheetName).Cells(13, 3).Text)
End Sub


Public Property Get InitDisp(ByVal i As Integer, ByVal j As Integer) As Double
    InitDisp = pInitDisp(i, j)
End Property

Public Property Let InitDisp(ByVal i As Integer, ByVal j As Integer, ByVal myVal As Double)
   pInitDisp(i, j) = myVal
End Property
