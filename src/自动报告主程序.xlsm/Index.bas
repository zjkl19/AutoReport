Attribute VB_Name = "Index"
Option Explicit

Private Sub Auto_Open()
    Worksheets("��ҳ").Activate
    'indexForm.Show vbModeless
End Sub

'ʹ��˵��
Private Sub Instructions_Click()
    MsgBox "1��ѡ��""Ӧ��""��ǩ������Ӧ��" _
    & vbCrLf & "2��ѡ��""�Ӷ�""��ǩ�������Ӷ�" _
    & vbCrLf & "3��ѡ��""����Word����""��ǩ������Word����" _
    & vbCrLf & "���֧��" & CStr(MAX_NWC) & "������", , "ʹ��˵��"
End Sub

Public Sub CheckUpdate()
Dim Url As String
Dim HttpReq As Object
Dim op1, op2 As Object    'ѡ��1,2

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
        MsgBox "��ǰ�������°汾���������"
    Else
        MsgBox "�����°汾��" & JSONParse("Version", HttpReq.ResponseText) _
        & vbCrLf & "�¹��ܣ�" & JSONParse("Feature", HttpReq.ResponseText) _
        & vbCrLf & "���ص�ַ��" & server & JSONParse("DownloadURL", HttpReq.ResponseText)
        
        Cells(10, 8) = "�°汾���ص�ַ��"
        Cells(10, 9) = server & JSONParse("DownloadURL", HttpReq.ResponseText)
    End If
    
    Set HttpReq = Nothing
    Exit Sub
    
    On Error GoTo MsgboxStatus:
MsgboxStatus:
    MsgBox HttpReq.Status
    Set HttpReq = Nothing
    
End Sub

'JSONPathΪ���ݷ���·��
'JSONStringΪJSON��ʽԴ����
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
