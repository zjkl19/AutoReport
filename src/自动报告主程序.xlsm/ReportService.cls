VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReportService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'用于处理word报告中排版等功能的类

'计算位移交叉引用位置（图表通用）
'groupIndex:组索引
'GlobalWC(1 to xx):位移全局工况定位数组
'nwcs:挠度工况数
'StrainGlobalWC(1 to xx):应变全局工况定位数组
'strainNWCS:应变工况数
'Offset:起点偏移（报告中有其它图表编号在此之前）
Public Function GetDispCrossReferenceItem(ByVal groupIndex As Integer, ByRef GlobalWC() As Integer, ByRef StrainGlobalWC() As Integer, ByVal strainNWCs As Integer, Optional Offset As Integer = 0) As Integer
    Dim i As Integer
    Dim result As Integer
    result = groupIndex  'result最小值为groupIndex，若有应变工况，还会增加
    For i = 1 To strainNWCs
        If StrainGlobalWC(i) < GlobalWC(groupIndex) Then
            result = result + 1
        End If
    Next
    GetDispCrossReferenceItem = result + Offset
End Function

'计算应变表交叉引用位置（图表通用）
'groupIndex:组索引
'GlobalWC(1 to xx):位移全局工况定位数组
'nwcs:挠度工况数
'StrainGlobalWC(1 to xx):应变全局工况定位数组
'strainNWCS:应变工况数
'Offset:起点偏移（报告中有其它图表编号在此之前）
Public Function GetStrainCrossReferenceItem(ByVal groupIndex As Integer, ByRef GlobalWC() As Integer, ByVal NWCs As Integer, ByRef StrainGlobalWC() As Integer, Optional Offset As Integer = 0) As Integer
    Dim i As Integer
    Dim result As Integer
    result = groupIndex  'result最小值为groupIndex，若有应变工况，还会增加
    For i = 1 To NWCs
        If StrainGlobalWC(groupIndex) >= GlobalWC(i) Then
            result = result + 1
        End If
    Next
     GetStrainCrossReferenceItem = result + Offset
End Function

'计算位移表交叉引用位置
'groupIndex:组索引
'GlobalWC(1 to xx):位移全局工况定位数组
'nwcs:挠度工况数
'StrainGlobalWC(1 to xx):应变全局工况定位数组
'strainNWCS:应变工况数
Public Function GetDispTbCrossReferenceItem(ByVal groupIndex As Integer, ByRef GlobalWC() As Integer, ByRef StrainGlobalWC() As Integer, ByVal strainNWCs As Integer) As Integer
    Dim i As Integer
    Dim result As Integer
    result = groupIndex  'result最小值为groupIndex，若有应变工况，还会增加
    For i = 1 To strainNWCs
        If StrainGlobalWC(i) < GlobalWC(groupIndex) Then
            result = result + 1
        End If
    Next
    GetDispTbCrossReferenceItem = result
End Function

'计算应变表交叉引用位置
'groupIndex:组索引
'GlobalWC(1 to xx):位移全局工况定位数组
'nwcs:挠度工况数
'StrainGlobalWC(1 to xx):应变全局工况定位数组
'strainNWCS:应变工况数
Public Function GetStrainTbCrossReferenceItem(ByVal groupIndex As Integer, ByRef GlobalWC() As Integer, ByVal NWCs As Integer, ByRef StrainGlobalWC() As Integer) As Integer
    Dim i As Integer
    Dim result As Integer
    result = groupIndex  'result最小值为groupIndex，若有应变工况，还会增加
    For i = 1 To NWCs
        If StrainGlobalWC(groupIndex) >= GlobalWC(i) Then
            result = result + 1
        End If
    Next
    GetStrainTbCrossReferenceItem = result
End Function

