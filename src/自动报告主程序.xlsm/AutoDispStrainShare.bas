Attribute VB_Name = "AutoDispStrainShare"
Public Const DispSheetName As String = "挠度"
Public Const StrainSheetName As String = "应变"
 
Public DispTheoryShapeObjArray(1 To MAX_NWC) As Shape    '存储各理论值图指针
Public DispTheoryShapeObjCounts As Integer


Public StrainTheoryShapeObjArray(1 To MAX_NWC) As Shape    '存储各理论值图指针
Public StrainTheoryShapeObjCounts As Integer

Public nPN    '各个工况对应中文名称

Public Const ExportPromteString As String = "导出完成！请再校核自动导出的结果，防止出错。"
Public Const FontName_KaiGB2312 As String = "楷体_GB2312"


'删除当前sheet中所有ChartObject
Sub DeleteAllCharts()
    Dim c As ChartObject
    For Each c In ActiveSheet.ChartObjects
        c.Delete
    Next
End Sub
