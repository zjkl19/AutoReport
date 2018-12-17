Attribute VB_Name = "Index"
Public Const DispSheetName As String = "挠度"

Private Sub Auto_Open()
    Worksheets("首页").Activate
    'indexForm.Show vbModeless
End Sub

'使用说明
Private Sub Instructions_Click()
    MsgBox "1、选择""应变""标签，计算应变" _
    & vbCrLf & "2、选择""挠度""标签，计算挠度" _
    & vbCrLf & "3、选择""生成Word报告""标签，导出Word报告" _
    & vbCrLf & "最多支持" & CStr(MAX_NWC) & "个工况", , "使用说明"
    
    
End Sub
