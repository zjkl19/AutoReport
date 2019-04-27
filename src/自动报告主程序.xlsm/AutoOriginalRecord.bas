Attribute VB_Name = "AutoOriginalRecord"
Public ProjectName As String
Public ProjectPlace As String    '项目地点
Public ProjectOverview As String    '工程概况
Public InspectionFoundation As String    '检测依据
Public InspectionContent As String    '检验内容
Public Client As String    '委托单位
Public ContractNo As String    '合同编号
Public InspectionTime As String    '检验时间

Private Const MAX_Instruments As Integer = 20    '最多仪器数量
Public instrumentCounts As Integer    '仪器数量
Public instrumentName(1 To MAX_Instruments) As String    '仪器名称
Public instrumentType(1 To MAX_Instruments) As String    '规格型号
Public InstrumentManagementNo(1 To MAX_Instruments) As String    '管理编号
Public InstrumentCalibrationData(1 To MAX_Instruments) As String

'从报告中提取项目名称、检验内容等关键信息
Public Sub GetContentFromReport()
    Dim reportFolder As String
    Dim reportName As String

    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    Dim tbl As Table
    
    Dim i As Integer
    Dim j As Integer
    Dim resultFlag As Boolean                    'True表示成功
    resultFlag = False

    reportFolder = "报告"
    '只查找1个后缀名为docx或者doc的文件
    reportName = Dir(ThisWorkbook.Path & "\" & reportFolder & "\*.docx")
    If reportName = "" Then
        reportName = Dir(ThisWorkbook.Path & "\" & reportFolder & "\*.doc")    'reportName = "酒店桥报告-改5.doc"
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

'获取仪器信息
'算法：查找第一个cell(1,1)中含有“仪器名称”的表格作为仪器信息的表格，根据该表格提取相关内容
'传入：InstrumentName：仪器名称，InstrumentType：仪器类型，InstrumentManagementNo：仪器管理编号
'返回：仪器数量
Public Function GetInstrumentInfo(ByRef doc As Word.Document, ByRef instrumentName() As String, ByRef instrumentType() As String, ByRef InstrumentManagementNo() As String)
    '----------------
    Dim folderName As String: folderName = "仪器信息数据库"
    Dim dbExcelName As String: dbExcelName = Cells(2, 2)   '"校准通知20190320.xls"
    
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
        If InStr(doc.Tables(i).Cell(1, 1).Range.Text, "仪器名称") > 0 Then
            Set tbl = doc.Tables(i)
            Exit For
        End If
    Next i
    'tbl.Rows.Count - 1:仪器数量
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
    Dim stringSearch As String   '用来查找的字符串
    'Set c = sheet.Range(Union(Range(Cells(4, 5), Cells(3000, 5)), Range(Cells(4, 10), Cells(3000, 10)))).Find(managementNo)
    
    If Mid(managementNo, 1, 4) = "02FB" Then
        stringSearch = Mid(managementNo, 1, 13)    '例：02FB050118002
    Else
        stringSearch = Mid(managementNo, 1, 9)    '例：(B)02-398
    End If
    
    Set c = sheet.UsedRange.Find(stringSearch)
    If Not c Is Nothing Then
        r = c.Row '返回行
        GetInstrumentCalibrationData = sheet.Cells(r, 8).value
        Exit Function
    End If
    'Dim i As Integer
    'For i = 107 To 3000
     '   If InStr(sheet.Cells(i, 5).value, managementNo) > 0 Or InStr(sheet.Cells(i, 10).value, managementNo) > 0 Then
      '      GetInstrumentCalibrationData = sheet.Cells(i, 8).value    '"2020-02-27" 参考格式
     '       Exit Function
     '   End If
    'Next i
NotFound:
    GetInstrumentCalibrationData = ""
End Function


'获取工程概况
'算法：截取"工程概况\r"至"桥梁结构布置"之间的内容
Public Function GetProjectOverviewFromReport(ByRef doc As Word.Document)
    On Error GoTo DefaultResult:
    Dim TargetRange As Object    'Dim TargetRange As range 类型和doc.Application.Selection.Range不匹配
    
    Dim mRegExp As Object       '正则表达式对象
    Dim mMatches As Object      '匹配字符串集合对象
    Dim mMatch As Object        '匹配字符串
    
    Dim matchResult As String
    
    doc.Application.Selection.WholeStory    '全选
    Set TargetRange = doc.Application.Selection.Range

    Set mRegExp = CreateObject("Vbscript.Regexp")    'New RegExp
    With mRegExp
        .Global = True                              'True表示匹配所有, False表示仅匹配第一个符合项
        .IgnoreCase = True                          'True表示不区分大小写, False表示区分大小写
        '(?<=工程概况)[\S\s]{2,1000}(，|。)(?=[\s\S]{1,20}图 1)
        '概况\n[\S\s]+[,，。]\n?[^\n]+图\s?1\S+。
        '工程概况\r[\S\s]{2,2000}桥梁结构布置
        '.Pattern = "概况\r[\S\s]+[,，。]\r?[^\r]+图\s?1\S*。"
        .Pattern = "概况\r[\S\s]+?[,，。]\r?[^\r]+?图\s?1\S*。"   '匹配字符模式  (工程概况)+([\s\S]*)+(桥梁结构布置) (?<=项目名称：).*?\r|(?<=工程名称：).*?\r
        Set mMatches = .Execute(TargetRange.Text)   '执行正则查找，返回所有匹配结果的集合，若未找到，则为空
        For Each mMatch In mMatches
            matchResult = CStr(mMatch.value)
        Next
    End With
    
    With mRegExp
        .Global = True
        .IgnoreCase = True
        .Pattern = "[,，。]\r?[^\r,，。]+图\s?1\S*。"
        matchResult = .Replace(matchResult, "。")
        
    End With
    
    Set mRegExp = Nothing
    Set mMatches = Nothing
    Set TargetRange = Nothing
    'Debug.Print matchResult
    matchResult = Replace(matchResult, "概况", "")
    matchResult = Replace(matchResult, vbCr, "")
    matchResult = Replace(matchResult, vbLf, "")
    GetProjectOverviewFromReport = matchResult
    Exit Function
DefaultResult:
    GetProjectOverviewFromReport = ""
End Function

'查找表格序号
'算法：cell(1,1)含有"委托单位"
Public Function FindTableNo(ByRef doc As Word.Document) As Long
    Dim i As Long: Dim tbl As Table
    For i = 0 To doc.Tables.Count - 1  'vba bug,i=1 to doc.tables.count无法使用doc.tables(i)寻址
        If InStr(doc.Tables(CLng(i + 1)).Cell(1, 1).Range.Text, "委托单位") > 0 Then
            FindTableNo = i + 1
            Exit Function
        End If
    Next
    FindTableNo = 0    '找不到返回0
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


'自动生成原始记录封面
Public Sub AutoCover()
    
    GetContentFromReport
    
    Dim originalRecordTemplateFolderName As String
    originalRecordTemplateFolderName = "原始记录模板"
    Dim coverTemplateFileName As String
    coverTemplateFileName = "03检测原始记录-封面.doc"

    
    Dim recordResultFolderName As String
    recordResultFolderName = "自动生成的原始记录"
    Dim coverResultFileName As String
    coverResultFileName = "03检测原始记录-封面.doc"

    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    Dim tbl As Table
    
    Dim i As Integer
    Dim j As Integer
    Dim resultFlag As Boolean                    'True表示成功
    resultFlag = False

    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\" & originalRecordTemplateFolderName & "\" & coverTemplateFileName, ThisWorkbook.Path & "\" & recordResultFolderName & "\" & coverResultFileName
      
    Set doc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & recordResultFolderName & "\" & coverResultFileName)
    
    doc.Tables(1).Cell(1, 1).Range.Text = vbCrLf & "工程名称：" & ProjectName
    doc.Tables(1).Cell(1, 1).Range.Font.name = "宋体"
    
    doc.Tables(1).Cell(2, 1).Range.Text = vbCrLf & "检测项目：" & InspectionContent
    doc.Tables(1).Cell(2, 1).Range.Font.name = "宋体"
    doc.Save
    
CloseWord:
    'wordApp.Documents.Close
    wordApp.Quit
    
    Set doc = Nothing
    Set wordApp = Nothing

End Sub

'打开封面
Public Sub OpenCover()
    Dim recordResultFolderName As String
    recordResultFolderName = "自动生成的原始记录"
    Dim coverResultFileName As String
    coverResultFileName = "03检测原始记录-封面.doc"
    Dim wordApp As Word.Application
    Set wordApp = CreateObject("Word.Application")
    wordApp.Documents.Open fileName:=ThisWorkbook.Path & "\" & recordResultFolderName & "\" & coverResultFileName, ReadOnly:=False
    wordApp.Visible = True
    wordApp.Activate
End Sub


'自动生成合同定期评审记录表
Public Sub AutoContractReview()

     GetContentFromReport
    
    Dim templateFolderName As String
    templateFolderName = "原始记录模板"
    Dim templateFileName As String
    templateFileName = "FJTCC-BG-0402D 合同定期评审记录表.doc"
      
    Dim resultFolderName As String
    resultFolderName = "自动生成的原始记录"
    Dim resultFileName As String
    resultFileName = "FJTCC-BG-0402D 合同定期评审记录表.doc"

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

'打开合同定期评审记录表
Public Sub OpenContractReview()
    Dim resultFolderName As String
    resultFolderName = "自动生成的原始记录"
    Dim resultFileName As String
    resultFileName = "FJTCC-BG-0402D 合同定期评审记录表.doc"
    Dim wordApp As Word.Application
    Set wordApp = CreateObject("Word.Application")
    wordApp.Documents.Open fileName:=ThisWorkbook.Path & "\" & resultFolderName & "\" & resultFileName, ReadOnly:=False
    wordApp.Visible = True
    wordApp.Activate

End Sub

'自动生成现场检测基本信息
Public Sub AutoInspectionBasicInfo()

    GetContentFromReport
    
    Dim instrumentStartRow  As Integer

    Dim templateFolderName As String
    templateFolderName = "原始记录模板"
    Dim templateFileName As String
    templateFileName = "现场检测基本信息.doc"
        
    Dim resultFolderName As String
    resultFolderName = "自动生成的原始记录"
    Dim resultFileName As String
    resultFileName = "现场检测基本信息.doc"

    Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    Dim tbl As Table
    
    Dim i As Integer: Dim j As Integer

    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    
    FileCopy ThisWorkbook.Path & "\" & templateFolderName & "\" & templateFileName, ThisWorkbook.Path & "\" & resultFolderName & "\" & resultFileName
      
    Set doc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & resultFolderName & "\" & resultFileName)
    
    '填写委托编号
    doc.Application.Selection.GoTo what:=wdGoToBookmark, name:="ContractNo"
    doc.Application.Selection.TypeText ContractNo
    
    '插入足够的行数，以填写仪器信息
    doc.Application.Selection.GoTo what:=wdGoToBookmark, name:="RowInsertStart"
    doc.Application.Selection.InsertRowsBelow NumRows:=instrumentCounts - 2  '已有2行
    
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

'打开现场检测基本信息
Public Sub OpenInspectionBasicInfo()
    Dim resultFolderName As String
    resultFolderName = "自动生成的原始记录"
    Dim resultFileName As String
    resultFileName = "现场检测基本信息.doc"
    Dim wordApp As Word.Application
    Set wordApp = CreateObject("Word.Application")
    wordApp.Documents.Open fileName:=ThisWorkbook.Path & "\" & resultFolderName & "\" & resultFileName, ReadOnly:=False
    wordApp.Visible = True
    wordApp.Activate

End Sub
