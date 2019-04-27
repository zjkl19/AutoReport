Attribute VB_Name = "AutoInstrument"
Option Explicit

Private Const SortOrderCol As Integer = 7    '仪器排序列
Private Const CharacterStringCol As Integer = 8
Private Const CharacterTypeCol As Integer = 9
Private Const ReportAssetNameCol As Integer = 10

Private Const SeriesNoCol As Integer = 12
Private Const ReportAssetNameCopyCol As Integer = 13
Private Const AssetNameCol As Integer = 14
Private Const AssetNoCol As Integer = 15
Private Const InstrumentTypeCol As Integer = 16
Private Const OldInstrumentNoCol As Integer = 17
Private Const CalibrationDataCol As Integer = 18
Private Const ManufNoCol As Integer = 19    '出厂编号
Private Const SelectionCol As Integer = 20
Private Const SelectionCountCol As Integer = 21

'报告数据：序号  仪器名称    型号规格    管理编号    校准有效期至
Private Const ReportSeriesNoCol As Integer = 23
Private Const ReportInstrumentNameCol As Integer = 24
Private Const ReportInstrumentTypeCol As Integer = 25
Private Const ReportManagementNoCol As Integer = 26
Private Const ReportCalibrationDataCol As Integer = 27
Private Const ReportOldInstrumentNoCol As Integer = 28
Private Const ReportManufNoCol As Integer = 29

Private Const MaxInstrumentCounts As Integer = 100
Private InstrumentChkBox(1 To MaxInstrumentCounts) As CheckBox

'打开荷载试验计算书报告
Public Sub OpenUpgradeInstrumentDbHelp()
    Dim wordApp As Word.Application
    Set wordApp = CreateObject("Word.Application")
    wordApp.Documents.Open fileName:=ThisWorkbook.Path & "\" & "如何更新仪器信息数据库（仅供参考）.docx", ReadOnly:=False
    wordApp.Visible = True
    wordApp.Activate
End Sub

Sub testcheckBox()
    Dim chkBox As CheckBox

    Dim startRow As Integer
    startRow = 2
    Dim c As Range
    '1200
    '1500
    For Each chkBox In ActiveSheet.CheckBoxes
        If chkBox.Left > 500 Then
            Debug.Print chkBox.TopLeftCell.Row
            Debug.Print chkBox.Left
        End If
    Next
End Sub

Sub excelAddCheckBox()
Dim rng As Range
Dim chk As CheckBox
For Each rng In Selection
    With rng
        Set chk = ActiveSheet.CheckBoxes.Add(.Left, .Top, .Width, .Height) '.Select
        chk.value = xlOff
        chk.caption = "选择"
    
    End With
Next
End Sub

Public Sub tt()
Dim firstRow As Integer
Dim lastRow As Integer
GetFirstAndLastCol "百分表", "(0～10)mm", firstRow, lastRow
    SortArea firstRow, lastRow, OldInstrumentNoCol, ReportAssetNameCopyCol, CalibrationDataCol
End Sub

'查找需要被排序的首行和尾行
'instrumentName：仪器名称
'instrumentType：仪器类型
'firstRow：首行
'lastRow：尾行
Public Sub GetFirstAndLastCol(ByVal instrumentName, ByVal instrumentType, ByRef firstRow As Integer, ByRef lastRow As Integer)
    Dim currRow As Integer
    currRow = 2
    Dim firstRowFound As Boolean: Dim lastRowFound As Boolean
    firstRowFound = False: lastRowFound = False
    firstRow = currRow
    lastRow = currRow
    
    Do
        If Cells(currRow, ReportAssetNameCopyCol) = instrumentName And Cells(currRow, InstrumentTypeCol) = instrumentType And firstRowFound = False Then
            firstRow = currRow
            firstRowFound = True
            GoTo NextLoop:
        End If
        If firstRowFound = True Then
            If Cells(currRow, ReportAssetNameCopyCol) <> instrumentName Or Cells(currRow, InstrumentTypeCol) <> instrumentType Then
                lastRow = currRow - 1
                lastRowFound = True
                Exit Do
            End If
        End If
NextLoop:
        currRow = currRow + 1
    Loop While firstRowFound = False Or lastRowFound = False
End Sub

'对指定区域根据旧系统仪器设备编号进行排序
'firstRow：首行
'lastRow：末行
'sortOnCol：排序依据的列
'sortLeftRangeCol：左边范围
'sortRightRangeCol：右边范围
Public Sub SortArea(ByVal firstRow As Integer, ByVal lastRow As Integer, ByVal sortOnCol As Integer, ByVal sortLeftRangeCol As Integer, ByVal sortRightRangeCol As Integer)

    With ActiveSheet.Sort
        With .SortFields
            .Clear
            .Add Key:=Range(Cells(firstRow, sortOnCol), Cells(lastRow, sortOnCol)), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:=""
        End With
        .Header = xlNo
        .Orientation = xlSortColumns
        .MatchCase = False
        .SortMethod = xlPinYin
        .SetRange rng:=Range(Cells(firstRow, sortLeftRangeCol), Cells(lastRow, sortRightRangeCol))
        .Apply
    End With
End Sub

'获取所选的仪器列表
Public Sub GetSelectedInstrument()
    ClearSelectedInstrumentList
    Const SelectionChkBoxLeftBoundry As Integer = 1000
    Const SelectionChkBoxRightBoundry As Integer = 1800
    
    Dim chkBox As CheckBox
    Dim currRow As Integer
    currRow = 2

    For Each chkBox In ActiveSheet.CheckBoxes
        If chkBox.value = xlOn And chkBox.Left > SelectionChkBoxLeftBoundry And chkBox.Left < SelectionChkBoxRightBoundry Then
            Cells(currRow, ReportSeriesNoCol) = currRow - 1
            Cells(currRow, ReportInstrumentNameCol) = Cells(chkBox.TopLeftCell.Row, ReportAssetNameCopyCol)
            Cells(currRow, ReportInstrumentTypeCol) = Cells(chkBox.TopLeftCell.Row, InstrumentTypeCol)
            Cells(currRow, ReportManagementNoCol) = Cells(chkBox.TopLeftCell.Row, AssetNoCol)
            Cells(currRow, ReportCalibrationDataCol) = CDate(Cells(chkBox.TopLeftCell.Row, CalibrationDataCol))
            Cells(currRow, ReportOldInstrumentNoCol) = Cells(chkBox.TopLeftCell.Row, OldInstrumentNoCol)
            Cells(currRow, ReportManufNoCol) = Cells(chkBox.TopLeftCell.Row, ManufNoCol)
            currRow = currRow + 1
        End If
    Next
    
End Sub

'清空已选仪器列表
Public Sub ClearSelectedInstrumentList()
    Dim currRow As Integer
    currRow = 2
    While Cells(currRow, SeriesNoCol) <> ""
        Cells(currRow, ReportSeriesNoCol) = ""
        Cells(currRow, ReportInstrumentNameCol) = ""
        Cells(currRow, ReportInstrumentTypeCol) = ""
        Cells(currRow, ReportManagementNoCol) = ""
        Cells(currRow, ReportCalibrationDataCol) = ""
        Cells(currRow, ReportOldInstrumentNoCol) = ""
        Cells(currRow, ReportManufNoCol) = ""
        currRow = currRow + 1
    Wend
End Sub

'清空列表
Public Sub ClearInstrumentList()
    Dim chkBox As CheckBox
    Dim currRow As Integer
    currRow = 2
    
    While Cells(currRow, SeriesNoCol) <> ""
        Cells(currRow, SeriesNoCol) = ""
        Cells(currRow, ReportAssetNameCopyCol) = ""
        Cells(currRow, AssetNameCol) = ""
        Cells(currRow, AssetNoCol) = ""
        Cells(currRow, InstrumentTypeCol) = ""
        Cells(currRow, OldInstrumentNoCol) = ""
        Cells(currRow, CalibrationDataCol) = ""
        Cells(currRow, ManufNoCol) = ""
        currRow = currRow + 1
    Wend
    
    Const SelectionChkBoxLeftBoundry As Integer = 1000
    Const SelectionChkBoxRightBoundry As Integer = 1500
    
    For Each chkBox In ActiveSheet.CheckBoxes
        If chkBox.Left > SelectionChkBoxLeftBoundry And chkBox.Left < SelectionChkBoxRightBoundry Then
            chkBox.Delete
        End If
    Next
    
End Sub

'列出可用仪器
Public Sub ListAvailableInstrument()
    '先清空列表，再列出可用仪器
    ClearInstrumentList
    
    Const AssetNameColInDb As Integer = 3
    Const AssetNoColInDb As Integer = 5
    Const CalibrationDataColInDb As Integer = 8
    Const InstrumentTypeColInDb As Integer = 9
    Const OldInstrumentNoColInDb As Integer = 10
    Const ManufNoColInDb As Integer = 13
    
    Const ChkBoxLeftBoundry As Integer = 100
    Const ChkBoxRightBoundry As Integer = 200
    
    Dim folderName As String: folderName = "仪器信息数据库"
    Dim dbExcelName As String: dbExcelName = Cells(2, 2)   '"校准通知20190320.xls"
    
    Dim currRow As Integer
    Dim chkBox As CheckBox
    Dim CharacterString As String    '特征字符串
    Dim CharacterType As String    '特征型号
    Dim ReportAssetName As String    '报告资产名称
    Dim CriticalDate As String    '临界校准日期
    CriticalDate = Cells(1, 2)
    
    Dim dataExcel As Excel.Application
    Dim Workbook As Excel.Workbook
    Dim sheet As Excel.Worksheet
    
    Set dataExcel = CreateObject("Excel.Application")
    Set Workbook = dataExcel.Workbooks.Open(ThisWorkbook.Path & "\" & folderName & "\" & dbExcelName)
    Set sheet = Workbook.Worksheets(1)

    Dim startRow As Integer
    startRow = 2
    Dim c As Range
    Dim firstAddress As String
    
    Dim selectionChk As CheckBox
    Dim i As Integer
    '对checkbox进行排序
    For i = 1 To MaxInstrumentCounts
           Set InstrumentChkBox(i) = Nothing
   Next
    Dim instrumentCounts As Integer
    instrumentCounts = 0
    For Each chkBox In ActiveSheet.CheckBoxes
        If chkBox = xlOn And chkBox.Left > ChkBoxLeftBoundry And chkBox.Left < ChkBoxRightBoundry Then
            Set InstrumentChkBox(Cells(chkBox.TopLeftCell.Row, SortOrderCol)) = chkBox
             instrumentCounts = instrumentCounts + 1
        End If
    Next
     
    
    For i = 1 To MaxInstrumentCounts
            
            Set chkBox = InstrumentChkBox(i)
            If chkBox Is Nothing Then
                GoTo nextFor:
            End If
            
            currRow = chkBox.TopLeftCell.Row
            CharacterString = Cells(currRow, CharacterStringCol)
            CharacterType = Cells(currRow, CharacterTypeCol)
            ReportAssetName = Cells(currRow, ReportAssetNameCol)
            With sheet.UsedRange
                Set c = .Find(CharacterString, LookIn:=xlValues)
                If Not c Is Nothing Then
                    firstAddress = c.Address
                    Do
                        If Len(Trim(sheet.Cells(c.Row, CalibrationDataColInDb).value)) > 0 And sheet.Cells(c.Row, CalibrationDataColInDb).value >= CDate(CriticalDate) Then '必须要有校准内容并且校准日期要大于临界校准日期
                            If CharacterType <> "" And sheet.Cells(c.Row, InstrumentTypeColInDb).value = CharacterType Then  '特征型号不为空，则必须同时满足特征型号的条件
                                Cells(startRow, SeriesNoCol) = startRow - 1
                                Cells(startRow, ReportAssetNameCopyCol) = ReportAssetName
                                Cells(startRow, AssetNameCol) = sheet.Cells(c.Row, AssetNameColInDb).value
                                Cells(startRow, AssetNoCol) = sheet.Cells(c.Row, AssetNoColInDb).value
                                Cells(startRow, InstrumentTypeCol) = sheet.Cells(c.Row, InstrumentTypeColInDb).value
                                Cells(startRow, OldInstrumentNoCol) = sheet.Cells(c.Row, OldInstrumentNoColInDb).value
                                
                                Cells(startRow, ManufNoCol).NumberFormatLocal = "@"    '设置单元格格式为文本
                                Cells(startRow, ManufNoCol) = CStr(sheet.Cells(c.Row, ManufNoColInDb).value)
                                
                                Cells(startRow, CalibrationDataCol) = sheet.Cells(c.Row, CalibrationDataColInDb).value
                                
                                With Cells(startRow, SelectionCol)
                                    Set selectionChk = ActiveSheet.CheckBoxes.Add(.Left, .Top, .Width, .Height)
                                End With
                                selectionChk.value = xlOff
                                selectionChk.caption = "选择"
                                
                                startRow = startRow + 1
                            End If
                            If CharacterType = "" Then
                                Cells(startRow, SeriesNoCol) = startRow - 1
                                Cells(startRow, ReportAssetNameCopyCol) = ReportAssetName
                                Cells(startRow, AssetNameCol) = sheet.Cells(c.Row, AssetNameColInDb).value
                                Cells(startRow, AssetNoCol) = sheet.Cells(c.Row, AssetNoColInDb).value
                                Cells(startRow, InstrumentTypeCol) = sheet.Cells(c.Row, InstrumentTypeColInDb).value
                                Cells(startRow, OldInstrumentNoCol) = sheet.Cells(c.Row, OldInstrumentNoColInDb).value
                                Cells(startRow, ManufNoCol) = sheet.Cells(c.Row, ManufNoColInDb).value
                                Cells(startRow, CalibrationDataCol) = sheet.Cells(c.Row, CalibrationDataColInDb).value
                                
                                With Cells(startRow, SelectionCol)
                                    Set selectionChk = ActiveSheet.CheckBoxes.Add(.Left, .Top, .Width, .Height)
                                End With
                                selectionChk.value = xlOff
                                selectionChk.caption = "选择"
                                
                                startRow = startRow + 1
                            End If

                        End If
                        Set c = .FindNext(c)
                        If c Is Nothing Then
                            GoTo DoneFinding
                        End If
                    Loop While c.Address <> firstAddress
               End If
DoneFinding:
                Set c = Nothing
            End With
nextFor:
    Next
    Dim firstRow As Integer
    Dim lastRow As Integer
   For Each chkBox In ActiveSheet.CheckBoxes
        If chkBox = xlOn And chkBox.Left > ChkBoxLeftBoundry And chkBox.Left < ChkBoxRightBoundry Then
            If Cells(chkBox.TopLeftCell.Row, CharacterStringCol) = "百分表" Then
                GetFirstAndLastCol "百分表", Cells(chkBox.TopLeftCell.Row, CharacterTypeCol), firstRow, lastRow
                SortArea firstRow, lastRow, OldInstrumentNoCol, ReportAssetNameCopyCol, CalibrationDataCol
            End If
        End If
    Next

    Workbook.Close
    dataExcel.Quit
    Set dataExcel = Nothing
    Set Workbook = Nothing
    Set sheet = Nothing

End Sub
