Attribute VB_Name = "AutoOriginalRecord"
Public ProjectName As String
Public ProjectPlace As String    '��Ŀ�ص�
Public ProjectOverview As String    '���̸ſ�
Public InspectionFoundation As String    '�������
Public InspectionContent As String    '��������
Public Client As String    'ί�е�λ
Public ContractNo As String    '��ͬ���
Public InspectionTime As String    '����ʱ��

Private Const MAX_Instruments As Integer = 20    '�����������
Public instrumentCounts As Integer    '��������
Public instrumentName(1 To MAX_Instruments) As String    '��������
Public instrumentType(1 To MAX_Instruments) As String    '����ͺ�
Public InstrumentManagementNo(1 To MAX_Instruments) As String    '������
Public InstrumentCalibrationData(1 To MAX_Instruments) As String

'�ӱ�������ȡ��Ŀ���ơ��������ݵȹؼ���Ϣ
Public Sub GetContentFromReport()
    Dim reportFolder As String
    Dim reportName As String

    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    Dim tbl As Table
    
    Dim i As Integer
    Dim j As Integer
    Dim resultFlag As Boolean                    'True��ʾ�ɹ�
    resultFlag = False

    reportFolder = "����"
    'ֻ����1����׺��Ϊdocx����doc���ļ�
    reportName = Dir(ThisWorkbook.Path & "\" & reportFolder & "\*.docx")
    If reportName = "" Then
        reportName = Dir(ThisWorkbook.Path & "\" & reportFolder & "\*.doc")    'reportName = "�Ƶ��ű���-��5.doc"
    End If
    
    'On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    
    'FileCopy ThisWorkbook.Path & "\" & reportTemplateFileName, ThisWorkbook.Path & "\AutoReportSource.docx"
    
    Set doc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & reportFolder & "\" & reportName)

    ProjectName = Replace(GetProjectNameFromReport(doc), "", "")
    ProjectName = Mid(ProjectName, 1, Len(ProjectName) - 1)
    
    InspectionContent = Replace(GetInspectionContentFromReport(doc), "", "")
    InspectionContent = Mid(InspectionContent, 1, Len(InspectionContent) - 1)
    
    Client = Replace(GetClientFromReport(doc), "", "")
    Client = Mid(Client, 1, Len(Client) - 1)
    
    ContractNo = Replace(GetContractNoFromReport(doc), "", "")
    ContractNo = Mid(ContractNo, 1, Len(ContractNo) - 1)
    
    InspectionTime = Replace(GetInspectionTimeFromReport(doc), "", "")
    InspectionTime = Mid(InspectionTime, 1, Len(InspectionTime) - 1)
    
    ProjectPlace = Replace(GetProjectPlaceFromReport(doc), "", "")
    ProjectPlace = Mid(ProjectPlace, 1, Len(ProjectPlace) - 1)
    
    InspectionFoundation = Replace(GetInspectionFoundationFromReport(doc), "", "")
    InspectionFoundation = Mid(InspectionFoundation, 1, Len(InspectionFoundation) - 1)
    
    ContractNo = Replace(GetContractNoFromReport(doc), "", "")
    ContractNo = Mid(ContractNo, 1, Len(ContractNo) - 1)
    
    
    ProjectOverview = Replace(GetProjectOverviewFromReport(doc), "", "")
    
    instrumentCounts = GetInstrumentInfo(doc, instrumentName, instrumentType, InstrumentManagementNo)
    
CloseWord:
    wordApp.Documents.Close
    wordApp.Quit
    
    Set doc = Nothing
    Set wordApp = Nothing
    'Set tbl = Nothing

End Sub

Public Sub AutoAllOriginalRecord()
    AutoCover
    AutoContractReview
    AutoInspectionBasicInfo
End Sub

'��ȡ������Ϣ
'�㷨�����ҵ�һ��cell(1,1)�к��С��������ơ��ı����Ϊ������Ϣ�ı�񣬸��ݸñ����ȡ�������
'���룺InstrumentName���������ƣ�InstrumentType���������ͣ�InstrumentManagementNo������������
'���أ���������
Public Function GetInstrumentInfo(ByRef doc As Word.Document, ByRef instrumentName() As String, ByRef instrumentType() As String, ByRef InstrumentManagementNo() As String)
    '----------------
    Dim folderName As String: folderName = "������Ϣ���ݿ�"
    Dim dbExcelName As String: dbExcelName = Cells(2, 2)   '"У׼֪ͨ20190320.xls"
    
    Dim dataExcel As Excel.Application
    Dim Workbook As Excel.Workbook
    Dim sheet As Excel.Worksheet
    
    Set dataExcel = CreateObject("Excel.Application")
    Set Workbook = dataExcel.Workbooks.Open(ThisWorkbook.Path & "\" & folderName & "\" & dbExcelName)
    Set sheet = Workbook.Worksheets(1)
     '----------------
    
    Dim tbl As Word.Table
    Dim i As Integer
    For i = 1 To doc.Tables.Count
        If InStr(doc.Tables(i).Cell(1, 1).Range.Text, "��������") > 0 Then
            Set tbl = doc.Tables(i)
            Exit For
        End If
    Next i
    'tbl.Rows.Count - 1:��������
    For i = 2 To tbl.Rows.Count
        instrumentName(i - 1) = Replace(tbl.Cell(i, 1).Range.Text, "", "")
        instrumentType(i - 1) = Replace(tbl.Cell(i, 2).Range.Text, "", "")
        InstrumentManagementNo(i - 1) = Replace(tbl.Cell(i, 3).Range.Text, "", "")
        InstrumentCalibrationData(i - 1) = GetInstrumentCalibrationData(sheet, InstrumentManagementNo(i - 1))
    Next i
    
    Workbook.Close
    dataExcel.Quit
    Set dataExcel = Nothing
    Set Workbook = Nothing
    Set sheet = Nothing
    
    GetInstrumentInfo = tbl.Rows.Count - 1
End Function

Public Function GetInstrumentCalibrationData(ByRef sheet As Excel.Worksheet, ByVal managementNo As String)
    'On Error GoTo NotFound:
    Dim c
    Dim r As Integer
    Dim stringSearch As String   '�������ҵ��ַ���
    'Set c = sheet.Range(Union(Range(Cells(4, 5), Cells(3000, 5)), Range(Cells(4, 10), Cells(3000, 10)))).Find(managementNo)
    
    If Mid(managementNo, 1, 4) = "02FB" Then
        stringSearch = Mid(managementNo, 1, 13)    '����02FB050118002
    Else
        stringSearch = Mid(managementNo, 1, 9)    '����(B)02-398
    End If
    
    Set c = sheet.UsedRange.Find(stringSearch)
    If Not c Is Nothing Then
        r = c.Row '������
        GetInstrumentCalibrationData = sheet.Cells(r, 8).value
        Exit Function
    End If
    'Dim i As Integer
    'For i = 107 To 3000
     '   If InStr(sheet.Cells(i, 5).value, managementNo) > 0 Or InStr(sheet.Cells(i, 10).value, managementNo) > 0 Then
      '      GetInstrumentCalibrationData = sheet.Cells(i, 8).value    '"2020-02-27" �ο���ʽ
     '       Exit Function
     '   End If
    'Next i
NotFound:
    GetInstrumentCalibrationData = ""
End Function


'��ȡ���̸ſ�
'�㷨����ȡ"���̸ſ�\r"��"�����ṹ����"֮�������
Public Function GetProjectOverviewFromReport(ByRef doc As Word.Document)
    On Error GoTo DefaultResult:
    Dim TargetRange As Object    'Dim TargetRange As range ���ͺ�doc.Application.Selection.Range��ƥ��
    
    Dim mRegExp As Object       '������ʽ����
    Dim mMatches As Object      'ƥ���ַ������϶���
    Dim mMatch As Object        'ƥ���ַ���
    
    Dim matchResult As String
    
    doc.Application.Selection.WholeStory    'ȫѡ
    Set TargetRange = doc.Application.Selection.Range

    Set mRegExp = CreateObject("Vbscript.Regexp")    'New RegExp
    With mRegExp
        .Global = True                              'True��ʾƥ������, False��ʾ��ƥ���һ��������
        .IgnoreCase = True                          'True��ʾ�����ִ�Сд, False��ʾ���ִ�Сд
        '(?<=���̸ſ�)[\S\s]{2,1000}(��|��)(?=[\s\S]{1,20}ͼ 1)
        '�ſ�\n[\S\s]+[,����]\n?[^\n]+ͼ\s?1\S+��
        '���̸ſ�\r[\S\s]{2,2000}�����ṹ����
        '.Pattern = "�ſ�\r[\S\s]+[,����]\r?[^\r]+ͼ\s?1\S*��"
        .Pattern = "�ſ�\r[\S\s]+?[,����]\r?[^\r]+?ͼ\s?1\S*��"   'ƥ���ַ�ģʽ  (���̸ſ�)+([\s\S]*)+(�����ṹ����) (?<=��Ŀ���ƣ�).*?\r|(?<=�������ƣ�).*?\r
        Set mMatches = .Execute(TargetRange.Text)   'ִ��������ң���������ƥ�����ļ��ϣ���δ�ҵ�����Ϊ��
        For Each mMatch In mMatches
            matchResult = CStr(mMatch.value)
        Next
    End With
    
    With mRegExp
        .Global = True
        .IgnoreCase = True
        .Pattern = "[,����]\r?[^\r,����]+ͼ\s?1\S*��"
        matchResult = .Replace(matchResult, "��")
        
    End With
    
    Set mRegExp = Nothing
    Set mMatches = Nothing
    Set TargetRange = Nothing
    'Debug.Print matchResult
    matchResult = Replace(matchResult, "�ſ�", "")
    matchResult = Replace(matchResult, vbCr, "")
    matchResult = Replace(matchResult, vbLf, "")
    GetProjectOverviewFromReport = matchResult
    Exit Function
DefaultResult:
    GetProjectOverviewFromReport = ""
End Function

'���ұ�����
'�㷨��cell(1,1)����"ί�е�λ"
Public Function FindTableNo(ByRef doc As Word.Document) As Long
    Dim i As Long: Dim tbl As Table
    For i = 0 To doc.Tables.Count - 1  'vba bug,i=1 to doc.tables.count�޷�ʹ��doc.tables(i)Ѱַ
        If InStr(doc.Tables(CLng(i + 1)).Cell(1, 1).Range.Text, "ί�е�λ") > 0 Then
            FindTableNo = i + 1
            Exit Function
        End If
    Next
    FindTableNo = 0    '�Ҳ�������0
End Function

Public Function GetProjectNameFromReport(ByRef doc As Word.Document)
    GetProjectNameFromReport = doc.Tables(FindTableNo(doc)).Cell(3, 2).Range.Text
End Function

Public Function GetContractNoFromReport(ByRef doc As Word.Document)
    GetContractNoFromReport = doc.Tables(FindTableNo(doc)).Cell(1, 5).Range.Text
End Function

Public Function GetInspectionContentFromReport(ByRef doc As Word.Document)
    GetInspectionContentFromReport = doc.Tables(FindTableNo(doc)).Cell(4, 2).Range.Text
End Function

Public Function GetClientFromReport(ByRef doc As Word.Document)
    GetClientFromReport = doc.Tables(FindTableNo(doc)).Cell(1, 3).Range.Text
End Function


Public Function GetInspectionTimeFromReport(ByRef doc As Word.Document)
    GetInspectionTimeFromReport = doc.Tables(FindTableNo(doc)).Cell(2, 5).Range.Text
End Function

Public Function GetProjectPlaceFromReport(ByRef doc As Word.Document)
    GetProjectPlaceFromReport = doc.Tables(FindTableNo(doc)).Cell(3, 4).Range.Text
End Function

Public Function GetInspectionFoundationFromReport(ByRef doc As Word.Document)
    GetInspectionFoundationFromReport = doc.Tables(FindTableNo(doc)).Cell(5, 2).Range.Text
End Function


'�Զ�����ԭʼ��¼����
Public Sub AutoCover()
    
    GetContentFromReport
    
    Dim originalRecordTemplateFolderName As String
    originalRecordTemplateFolderName = "ԭʼ��¼ģ��"
    Dim coverTemplateFileName As String
    coverTemplateFileName = "03���ԭʼ��¼-����.doc"

    
    Dim recordResultFolderName As String
    recordResultFolderName = "�Զ����ɵ�ԭʼ��¼"
    Dim coverResultFileName As String
    coverResultFileName = "03���ԭʼ��¼-����.doc"

    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    Dim tbl As Table
    
    Dim i As Integer
    Dim j As Integer
    Dim resultFlag As Boolean                    'True��ʾ�ɹ�
    resultFlag = False

    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\" & originalRecordTemplateFolderName & "\" & coverTemplateFileName, ThisWorkbook.Path & "\" & recordResultFolderName & "\" & coverResultFileName
      
    Set doc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & recordResultFolderName & "\" & coverResultFileName)
    
    doc.Tables(1).Cell(1, 1).Range.Text = vbCrLf & "�������ƣ�" & ProjectName
    doc.Tables(1).Cell(1, 1).Range.Font.name = "����"
    
    doc.Tables(1).Cell(2, 1).Range.Text = vbCrLf & "�����Ŀ��" & InspectionContent
    doc.Tables(1).Cell(2, 1).Range.Font.name = "����"
    doc.Save
    
CloseWord:
    'wordApp.Documents.Close
    wordApp.Quit
    
    Set doc = Nothing
    Set wordApp = Nothing

End Sub

'�򿪷���
Public Sub OpenCover()
    Dim recordResultFolderName As String
    recordResultFolderName = "�Զ����ɵ�ԭʼ��¼"
    Dim coverResultFileName As String
    coverResultFileName = "03���ԭʼ��¼-����.doc"
    Dim wordApp As Word.Application
    Set wordApp = CreateObject("Word.Application")
    wordApp.Documents.Open fileName:=ThisWorkbook.Path & "\" & recordResultFolderName & "\" & coverResultFileName, ReadOnly:=False
    wordApp.Visible = True
    wordApp.Activate
End Sub


'�Զ����ɺ�ͬ���������¼��
Public Sub AutoContractReview()

     GetContentFromReport
    
    Dim templateFolderName As String
    templateFolderName = "ԭʼ��¼ģ��"
    Dim templateFileName As String
    templateFileName = "FJTCC-BG-0402D ��ͬ���������¼��.doc"
      
    Dim resultFolderName As String
    resultFolderName = "�Զ����ɵ�ԭʼ��¼"
    Dim resultFileName As String
    resultFileName = "FJTCC-BG-0402D ��ͬ���������¼��.doc"

    Dim wordApp As Word.Application
    Dim doc As Word.Document
    
    Dim i As Integer
    Dim j As Integer

    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\" & templateFolderName & "\" & templateFileName, ThisWorkbook.Path & "\" & resultFolderName & "\" & resultFileName
      
    Set doc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & resultFolderName & "\" & resultFileName)
    
    doc.Tables(1).Cell(1, 2).Range.Text = ContractNo
    doc.Tables(1).Cell(1, 4).Range.Text = Client
    doc.Tables(1).Cell(2, 4).Range.Text = ProjectName
    
    doc.Save
    
CloseWord:
    'wordApp.Documents.Close
    wordApp.Quit
    
    Set doc = Nothing
    Set wordApp = Nothing

End Sub

'�򿪺�ͬ���������¼��
Public Sub OpenContractReview()
    Dim resultFolderName As String
    resultFolderName = "�Զ����ɵ�ԭʼ��¼"
    Dim resultFileName As String
    resultFileName = "FJTCC-BG-0402D ��ͬ���������¼��.doc"
    Dim wordApp As Word.Application
    Set wordApp = CreateObject("Word.Application")
    wordApp.Documents.Open fileName:=ThisWorkbook.Path & "\" & resultFolderName & "\" & resultFileName, ReadOnly:=False
    wordApp.Visible = True
    wordApp.Activate

End Sub

'�Զ������ֳ���������Ϣ
Public Sub AutoInspectionBasicInfo()

    GetContentFromReport
    
    Dim instrumentStartRow  As Integer

    Dim templateFolderName As String
    templateFolderName = "ԭʼ��¼ģ��"
    Dim templateFileName As String
    templateFileName = "�ֳ���������Ϣ.doc"
        
    Dim resultFolderName As String
    resultFolderName = "�Զ����ɵ�ԭʼ��¼"
    Dim resultFileName As String
    resultFileName = "�ֳ���������Ϣ.doc"

    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    Dim tbl As Table
    
    Dim i As Integer: Dim j As Integer

    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\" & templateFolderName & "\" & templateFileName, ThisWorkbook.Path & "\" & resultFolderName & "\" & resultFileName
      
    Set doc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & resultFolderName & "\" & resultFileName)
    
    '��дί�б��
    doc.Application.Selection.GoTo what:=wdGoToBookmark, name:="ContractNo"
    doc.Application.Selection.TypeText ContractNo
    
    '�����㹻������������д������Ϣ
    doc.Application.Selection.GoTo what:=wdGoToBookmark, name:="RowInsertStart"
    doc.Application.Selection.InsertRowsBelow NumRows:=instrumentCounts - 2  '����2��
    
    instrumentStartRow = 9
    For i = 1 To instrumentCounts
        doc.Tables(1).Cell(instrumentStartRow - 1 + i, 2).Range.Text = instrumentName(i)
        doc.Tables(1).Cell(instrumentStartRow - 1 + i, 3).Range.Text = instrumentType(i)
        doc.Tables(1).Cell(instrumentStartRow - 1 + i, 4).Range.Text = InstrumentManagementNo(i)
        doc.Tables(1).Cell(instrumentStartRow - 1 + i, 5).Range.Text = InstrumentCalibrationData(i)
    Next i
    
    doc.Tables(1).Cell(1, 2).Range.Text = Client
    doc.Tables(1).Cell(1, 4).Range.Text = InspectionTime
    doc.Tables(1).Cell(2, 2).Range.Text = ProjectName
    doc.Tables(1).Cell(2, 4).Range.Text = ProjectPlace
    doc.Tables(1).Cell(5, 2).Range.Text = ProjectOverview
    doc.Tables(1).Cell(6, 2).Range.Text = InspectionContent
    doc.Tables(1).Cell(7, 2).Range.Text = InspectionFoundation
    
    doc.Save
    
CloseWord:
    'wordApp.Documents.Close
    wordApp.Quit
    
    Set doc = Nothing
    Set wordApp = Nothing

End Sub

'���ֳ���������Ϣ
Public Sub OpenInspectionBasicInfo()
    Dim resultFolderName As String
    resultFolderName = "�Զ����ɵ�ԭʼ��¼"
    Dim resultFileName As String
    resultFileName = "�ֳ���������Ϣ.doc"
    Dim wordApp As Word.Application
    Set wordApp = CreateObject("Word.Application")
    wordApp.Documents.Open fileName:=ThisWorkbook.Path & "\" & resultFolderName & "\" & resultFileName, ReadOnly:=False
    wordApp.Visible = True
    wordApp.Activate

End Sub
