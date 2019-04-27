Attribute VB_Name = "Index"
Option Explicit

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

Public Sub CheckUpdate()
Dim Url As String
Dim HttpReq As Object
Dim op1, op2 As Object    '选项1,2

Dim server As String
Dim serverUpdateFile As String

If ActiveSheet.OptionButtons("op1") = xlOn Then
    server = "http://192.168.12.11:8300/"  '127.0.0.1
ElseIf ActiveSheet.OptionButtons("op2") = xlOn Then
    server = "http://218.66.5.89:8300/"
End If

serverUpdateFile = "AutoReportUpdate.txt"

Set HttpReq = CreateObject("MSXML2.ServerXMLHTTP")

    Url = server & serverUpdateFile
    
    HttpReq.Open "get", Url, False
    HttpReq.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    HttpReq.send
    
    If ActiveSheet.Labels("Version").Text = JSONParse("Version", HttpReq.ResponseText) Then
        MsgBox "当前已是最新版本，无须更新"
    Else
        MsgBox "现有新版本：" & JSONParse("Version", HttpReq.ResponseText) _
        & vbCrLf & "新功能：" & JSONParse("Feature", HttpReq.ResponseText) _
        & vbCrLf & "下载地址：" & server & JSONParse("DownloadURL", HttpReq.ResponseText)
        
        Cells(10, 8) = "新版本下载地址："
        Cells(10, 9) = server & JSONParse("DownloadURL", HttpReq.ResponseText)
    End If
    
    Set HttpReq = Nothing
    Exit Sub
    
    On Error GoTo MsgboxStatus:
MsgboxStatus:
    MsgBox HttpReq.Status
    Set HttpReq = Nothing
    
End Sub

'JSONPath为数据访问路径
'JSONString为JSON格式源数据
Public Function JSONParse(ByVal JSONPath As String, ByVal JSONString As String) As Variant
    Dim JSON As Object
    Set JSON = CreateObject("MSScriptControl.ScriptControl")
    JSON.Language = "JScript"
    JSONParse = JSON.Eval("JSON=" & JSONString & ";JSON." & JSONPath & ";")
    Set JSON = Nothing
End Function

Private Sub RemoveBrokenReferences(ByVal Book As Workbook)

'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dim oRefS As Object, oRef As Object
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Set oRefS = Book.VBProject.References
For Each oRef In oRefS
    If oRef.IsBroken Then
        Call oRefS.Remove(oRef)
    End If
Next

End Sub
