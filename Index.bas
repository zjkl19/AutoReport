Attribute VB_Name = "Index"
Public Const DispSheetName As String = "�Ӷ�"

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
