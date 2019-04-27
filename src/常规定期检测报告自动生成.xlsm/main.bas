Attribute VB_Name = "main"
Option Explicit

Public Const InspectionSheetName = "常规定期检测"

Public Const BridgeDeck As String = "桥面系"
'Word中对应书签BridgeDeckStart
Public Const SuperStructure As String = "上部结构"
Public Const SubStructure As String = "下部结构"

Private Const MaxDescriptions As Integer = 1000   '最大可描述数量
Private Const MaxPictures As Integer = 50   '每个描述最大照片数量

Private Const BridgePartColumn As Integer = 11: Private Const PositionColumn As Integer = 12
Private Const ComponentTypeColumn As Integer = 13: Private Const DamageTypeColumn As Integer = 14
Private Const DamageDescriptionColumn As Integer = 15: Private Const PictureDescriptionColumn As Integer = 16
Private Const PictureNoColumn As Integer = 17

Public BridgeDeckInspectionResultTable(1 To MaxDescriptions, 1 To 10) As String
Public BridgeDeckPictures(1 To MaxDescriptions, 1 To MaxPictures) As String    '桥面系每张照片的尾号
Public BridgeDeckPicturesFrontToEndIndex(1 To MaxDescriptions, 1 To 2) As Integer    '桥面系每个描述照片首尾编号

Public BridgeDeckDescriptionCounts As Integer    '桥面系描述数量:
Public BridgeDeckPictureCounts As Integer    '桥面系照片数量
Public BridgeDeckSpanCounts As Integer    '桥面系跨数

Public SuperStructureInspectionResultTable(1 To MaxDescriptions, 1 To 10) As String
Public SuperStructurePictures(1 To MaxDescriptions, 1 To MaxPictures) As String    '上部结构每张照片的尾号
Public SuperStructurePicturesFrontToEndIndex(1 To MaxDescriptions, 1 To 2) As Integer    '上部结构每个描述照片首尾编号
Public SuperStructureDescriptionCounts As Integer    '上部结构描述数量
Public SuperStructurePictureCounts As Integer    '上部结构照片数量
Public SuperStructureSpanCounts As Integer    '上部结构跨数

Public SubStructureInspectionResultTable(1 To MaxDescriptions, 1 To 10) As String
Public SubStructurePictures(1 To MaxDescriptions, 1 To MaxPictures) As String
Public SubStructurePicturesFrontToEndIndex(1 To MaxDescriptions, 1 To 2) As Integer
Public SubStructureDescriptionCounts As Integer
Public SubStructurePictureCounts As Integer
Public SubStructureSpanCounts As Integer

'获取检测结果基本信息
Public Sub GetInspectionResult()
    Dim startRow As Integer: startRow = 2
    Dim currRow As Integer: currRow = startRow
    Dim i As Integer: Dim j As Integer: i = 1
    Dim rs As ReportService
    
    BridgeDeckDescriptionCounts = 0: BridgeDeckPictureCounts = 0: BridgeDeckSpanCounts = 1
    Dim pictureNoOut As String
    Dim singlePictureCounts As Integer
    Dim pictures
    
    '初始化检测结果变量
'sheetName：标签名
'structureType：结构类型：桥面系，上部结构，下部结构
'InspectionResultTable()：2维检测结果数组
'PicturesFrontToEndIndex()：照片首尾索引
'DescriptionCounts：描述数量统计值
'PictureCounts：照片数量统计值
'SpanCounts：跨数统计值
'Public Sub InitInspectionResult(ByVal sheetName As String, ByVal structureType As String, ByRef InspectionResultTable() As String, ByRef PicturesFrontToEndIndex() As Integer _
'    , ByRef DescriptionCounts As Integer, ByRef PictureCounts As Integer, ByRef SpanCounts As Integer)
    Set rs = New ReportService
    '读入下部结构信息
    rs.InitInspectionResult sheetName:=InspectionSheetName, structureType:=SubStructure, InspectionResultTable:=SubStructureInspectionResultTable _
    , StorePictures:=SubStructurePictures, PicturesFrontToEndIndex:=SubStructurePicturesFrontToEndIndex, DescriptionCounts:=SubStructureDescriptionCounts _
    , PictureCounts:=SubStructurePictureCounts, SpanCounts:=SubStructureSpanCounts

    '读入上部结构信息
    rs.InitInspectionResult sheetName:=InspectionSheetName, structureType:=SuperStructure, InspectionResultTable:=SuperStructureInspectionResultTable _
    , StorePictures:=SuperStructurePictures, PicturesFrontToEndIndex:=SuperStructurePicturesFrontToEndIndex, DescriptionCounts:=SuperStructureDescriptionCounts _
    , PictureCounts:=SuperStructurePictureCounts, SpanCounts:=SuperStructureSpanCounts

    '读入桥面系信息
    While Cells(currRow, 11) = BridgeDeck    '只要有"桥面系"文字就继续读取下一行
        BridgeDeckInspectionResultTable(i, 1) = Cells(currRow, BridgePartColumn): BridgeDeckInspectionResultTable(i, 2) = Cells(currRow, PositionColumn)
        BridgeDeckInspectionResultTable(i, 3) = Cells(currRow, ComponentTypeColumn): BridgeDeckInspectionResultTable(i, 4) = Cells(currRow, DamageTypeColumn)
        BridgeDeckInspectionResultTable(i, 5) = Cells(currRow, DamageDescriptionColumn): BridgeDeckInspectionResultTable(i, 6) = Cells(currRow, PictureDescriptionColumn)
        BridgeDeckInspectionResultTable(i, 7) = CStr(Cells(currRow, PictureNoColumn))
        singlePictureCounts = rs.TranslatePictureNo(BridgeDeckInspectionResultTable(i, 7), pictureNoOut)
        BridgeDeckInspectionResultTable(i, 8) = CStr(pictureNoOut)
        BridgeDeckInspectionResultTable(i, 9) = UBound(Split(CStr(pictureNoOut), ",")) + 1    '相片数量
        
        'TODO：抽取计算跨数的方法
        '桥跨默认只有1行，若从第2行起，改行桥跨名与上一行不同，则+1
        If i > 1 Then
            If BridgeDeckInspectionResultTable(i, 2) <> BridgeDeckInspectionResultTable(i - 1, 2) Then
                BridgeDeckSpanCounts = BridgeDeckSpanCounts + 1
                Debug.Print BridgeDeckSpanCounts
            End If
        End If
        'Debug.Print BridgeDeckInspectionResultTable(i, 9)
        
        If pictureNoOut <> "" Then  '如果图片数量不为0
            pictures = Split(CStr(pictureNoOut), ",")
                For j = 1 To CInt(BridgeDeckInspectionResultTable(i, 9))    '依次存储各个相片尾号
                    BridgeDeckPictures(i, j) = CStr(pictures(j - 1))

                    'Debug.Print BridgeDeckPictures(i, j)
                Next j
        End If
        
        If pictureNoOut = "" Then  '图片数量为0
            BridgeDeckPicturesFrontToEndIndex(i, 1) = 0
            BridgeDeckPicturesFrontToEndIndex(i, 2) = 0
        Else
            If i = 1 Then  '如果是第1张
                BridgeDeckPicturesFrontToEndIndex(i, 1) = 1
                BridgeDeckPicturesFrontToEndIndex(i, 2) = 1 + CInt(BridgeDeckInspectionResultTable(i, 9)) - 1
            Else
                BridgeDeckPicturesFrontToEndIndex(i, 1) = rs.SumArray(BridgeDeckInspectionResultTable, 9, i - 1) + 1
                BridgeDeckPicturesFrontToEndIndex(i, 2) = rs.SumArray(BridgeDeckInspectionResultTable, 9, i - 1) + CInt(BridgeDeckInspectionResultTable(i, 9))
            End If
        End If
        'Debug.Print BridgeDeckPicturesFrontToEndIndex(i, 1)
        'Debug.Print BridgeDeckPicturesFrontToEndIndex(i, 2)

        
        BridgeDeckPictureCounts = BridgeDeckPictureCounts + singlePictureCounts     '算法：UBound(数组名)-LBound(数组名)+1
        
        i = i + 1: currRow = currRow + 1: BridgeDeckDescriptionCounts = BridgeDeckDescriptionCounts + 1
    Wend
    Set rs = Nothing
    'Debug.Print BridgeDeckPictureCounts
End Sub

Public Sub t()
    Dim pictureNoIn As String
    pictureNoIn = "1,2,3,4,5"
    Dim pictureNoOut
    pictureNoOut = Split(pictureNoIn, ",")
    Debug.Print pictureNoOut(0)
End Sub
