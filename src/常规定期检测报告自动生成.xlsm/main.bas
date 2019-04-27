Attribute VB_Name = "main"
Option Explicit

Public Const InspectionSheetName = "���涨�ڼ��"

Public Const BridgeDeck As String = "����ϵ"
'Word�ж�Ӧ��ǩBridgeDeckStart
Public Const SuperStructure As String = "�ϲ��ṹ"
Public Const SubStructure As String = "�²��ṹ"

Private Const MaxDescriptions As Integer = 1000   '������������
Private Const MaxPictures As Integer = 50   'ÿ�����������Ƭ����

Private Const BridgePartColumn As Integer = 11: Private Const PositionColumn As Integer = 12
Private Const ComponentTypeColumn As Integer = 13: Private Const DamageTypeColumn As Integer = 14
Private Const DamageDescriptionColumn As Integer = 15: Private Const PictureDescriptionColumn As Integer = 16
Private Const PictureNoColumn As Integer = 17

Public BridgeDeckInspectionResultTable(1 To MaxDescriptions, 1 To 10) As String
Public BridgeDeckPictures(1 To MaxDescriptions, 1 To MaxPictures) As String    '����ϵÿ����Ƭ��β��
Public BridgeDeckPicturesFrontToEndIndex(1 To MaxDescriptions, 1 To 2) As Integer    '����ϵÿ��������Ƭ��β���

Public BridgeDeckDescriptionCounts As Integer    '����ϵ��������:
Public BridgeDeckPictureCounts As Integer    '����ϵ��Ƭ����
Public BridgeDeckSpanCounts As Integer    '����ϵ����

Public SuperStructureInspectionResultTable(1 To MaxDescriptions, 1 To 10) As String
Public SuperStructurePictures(1 To MaxDescriptions, 1 To MaxPictures) As String    '�ϲ��ṹÿ����Ƭ��β��
Public SuperStructurePicturesFrontToEndIndex(1 To MaxDescriptions, 1 To 2) As Integer    '�ϲ��ṹÿ��������Ƭ��β���
Public SuperStructureDescriptionCounts As Integer    '�ϲ��ṹ��������
Public SuperStructurePictureCounts As Integer    '�ϲ��ṹ��Ƭ����
Public SuperStructureSpanCounts As Integer    '�ϲ��ṹ����

Public SubStructureInspectionResultTable(1 To MaxDescriptions, 1 To 10) As String
Public SubStructurePictures(1 To MaxDescriptions, 1 To MaxPictures) As String
Public SubStructurePicturesFrontToEndIndex(1 To MaxDescriptions, 1 To 2) As Integer
Public SubStructureDescriptionCounts As Integer
Public SubStructurePictureCounts As Integer
Public SubStructureSpanCounts As Integer

'��ȡ�����������Ϣ
Public Sub GetInspectionResult()
    Dim startRow As Integer: startRow = 2
    Dim currRow As Integer: currRow = startRow
    Dim i As Integer: Dim j As Integer: i = 1
    Dim rs As ReportService
    
    BridgeDeckDescriptionCounts = 0: BridgeDeckPictureCounts = 0: BridgeDeckSpanCounts = 1
    Dim pictureNoOut As String
    Dim singlePictureCounts As Integer
    Dim pictures
    
    '��ʼ�����������
'sheetName����ǩ��
'structureType���ṹ���ͣ�����ϵ���ϲ��ṹ���²��ṹ
'InspectionResultTable()��2ά���������
'PicturesFrontToEndIndex()����Ƭ��β����
'DescriptionCounts����������ͳ��ֵ
'PictureCounts����Ƭ����ͳ��ֵ
'SpanCounts������ͳ��ֵ
'Public Sub InitInspectionResult(ByVal sheetName As String, ByVal structureType As String, ByRef InspectionResultTable() As String, ByRef PicturesFrontToEndIndex() As Integer _
'    , ByRef DescriptionCounts As Integer, ByRef PictureCounts As Integer, ByRef SpanCounts As Integer)
    Set rs = New ReportService
    '�����²��ṹ��Ϣ
    rs.InitInspectionResult sheetName:=InspectionSheetName, structureType:=SubStructure, InspectionResultTable:=SubStructureInspectionResultTable _
    , StorePictures:=SubStructurePictures, PicturesFrontToEndIndex:=SubStructurePicturesFrontToEndIndex, DescriptionCounts:=SubStructureDescriptionCounts _
    , PictureCounts:=SubStructurePictureCounts, SpanCounts:=SubStructureSpanCounts

    '�����ϲ��ṹ��Ϣ
    rs.InitInspectionResult sheetName:=InspectionSheetName, structureType:=SuperStructure, InspectionResultTable:=SuperStructureInspectionResultTable _
    , StorePictures:=SuperStructurePictures, PicturesFrontToEndIndex:=SuperStructurePicturesFrontToEndIndex, DescriptionCounts:=SuperStructureDescriptionCounts _
    , PictureCounts:=SuperStructurePictureCounts, SpanCounts:=SuperStructureSpanCounts

    '��������ϵ��Ϣ
    While Cells(currRow, 11) = BridgeDeck    'ֻҪ��"����ϵ"���־ͼ�����ȡ��һ��
        BridgeDeckInspectionResultTable(i, 1) = Cells(currRow, BridgePartColumn): BridgeDeckInspectionResultTable(i, 2) = Cells(currRow, PositionColumn)
        BridgeDeckInspectionResultTable(i, 3) = Cells(currRow, ComponentTypeColumn): BridgeDeckInspectionResultTable(i, 4) = Cells(currRow, DamageTypeColumn)
        BridgeDeckInspectionResultTable(i, 5) = Cells(currRow, DamageDescriptionColumn): BridgeDeckInspectionResultTable(i, 6) = Cells(currRow, PictureDescriptionColumn)
        BridgeDeckInspectionResultTable(i, 7) = CStr(Cells(currRow, PictureNoColumn))
        singlePictureCounts = rs.TranslatePictureNo(BridgeDeckInspectionResultTable(i, 7), pictureNoOut)
        BridgeDeckInspectionResultTable(i, 8) = CStr(pictureNoOut)
        BridgeDeckInspectionResultTable(i, 9) = UBound(Split(CStr(pictureNoOut), ",")) + 1    '��Ƭ����
        
        'TODO����ȡ��������ķ���
        '�ſ�Ĭ��ֻ��1�У����ӵ�2���𣬸����ſ�������һ�в�ͬ����+1
        If i > 1 Then
            If BridgeDeckInspectionResultTable(i, 2) <> BridgeDeckInspectionResultTable(i - 1, 2) Then
                BridgeDeckSpanCounts = BridgeDeckSpanCounts + 1
                Debug.Print BridgeDeckSpanCounts
            End If
        End If
        'Debug.Print BridgeDeckInspectionResultTable(i, 9)
        
        If pictureNoOut <> "" Then  '���ͼƬ������Ϊ0
            pictures = Split(CStr(pictureNoOut), ",")
                For j = 1 To CInt(BridgeDeckInspectionResultTable(i, 9))    '���δ洢������Ƭβ��
                    BridgeDeckPictures(i, j) = CStr(pictures(j - 1))

                    'Debug.Print BridgeDeckPictures(i, j)
                Next j
        End If
        
        If pictureNoOut = "" Then  'ͼƬ����Ϊ0
            BridgeDeckPicturesFrontToEndIndex(i, 1) = 0
            BridgeDeckPicturesFrontToEndIndex(i, 2) = 0
        Else
            If i = 1 Then  '����ǵ�1��
                BridgeDeckPicturesFrontToEndIndex(i, 1) = 1
                BridgeDeckPicturesFrontToEndIndex(i, 2) = 1 + CInt(BridgeDeckInspectionResultTable(i, 9)) - 1
            Else
                BridgeDeckPicturesFrontToEndIndex(i, 1) = rs.SumArray(BridgeDeckInspectionResultTable, 9, i - 1) + 1
                BridgeDeckPicturesFrontToEndIndex(i, 2) = rs.SumArray(BridgeDeckInspectionResultTable, 9, i - 1) + CInt(BridgeDeckInspectionResultTable(i, 9))
            End If
        End If
        'Debug.Print BridgeDeckPicturesFrontToEndIndex(i, 1)
        'Debug.Print BridgeDeckPicturesFrontToEndIndex(i, 2)

        
        BridgeDeckPictureCounts = BridgeDeckPictureCounts + singlePictureCounts     '�㷨��UBound(������)-LBound(������)+1
        
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
