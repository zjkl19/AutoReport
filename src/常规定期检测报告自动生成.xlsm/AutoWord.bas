Attribute VB_Name = "AutoWord"
Option Explicit

Public Const BridgePartColumn As Integer = 11: Public Const PositionColumn As Integer = 12
Public Const ComponentTypeColumn As Integer = 13: Public Const DamageTypeColumn As Integer = 14
Public Const DamageDescriptionColumn As Integer = 15: Public Const PictureDescriptionColumn As Integer = 16
Public Const PictureNoColumn As Integer = 17

Public Const PositionIndex As Integer = 2: Public Const ComponentTypeIndex As Integer = 3
Public Const DamageTypeIndex As Integer = 4: Public Const DamageDescriptionIndex As Integer = 5
Public Const PictureDescriptionIndex As Integer = 6: Public Const PictureNoIndex As Integer = 7
Public Const PictureNoOutIndex As Integer = 8: Public Const PictureCountsIndex As Integer = 9

Public Sub OpenReport()
    Dim rs As ReportService
    Dim fileName As String
    Set rs = New ReportService
    fileName = ThisWorkbook.Path & "\�Զ����ɵĳ��涨�ڼ�ⱨ��\�Զ����ɵ��������涨�ڼ�ⱨ��.docx"
    rs.OpenReport fileName
    Set rs = Nothing
End Sub

'���ȼ��㣬������Word����
Public Sub AutoReport()
    
    GetInspectionResult    '��ȡ�����
    
    Const tableColumnCounts As Integer = 6  '���������δ�ϲ����ⵥԪ��
    
    Dim templateFolderName As String: Dim templateFileName As String
    templateFolderName = "���涨�ڼ�ⱨ��ģ��": templateFileName = "�������涨�ڼ�ⱨ��ģ��.docx"
    Dim resultFolderName As String: Dim resultFileName As String
    resultFolderName = "�Զ����ɵĳ��涨�ڼ�ⱨ��": resultFileName = "�Զ����ɵ��������涨�ڼ�ⱨ��.docx"
    Dim pictureFolderName As String: pictureFolderName = "���涨�ڼ����Ƭ"
    Dim currPictureName As String    '��ǰ��Ƭ����
    
    Dim wordApp As Word.Application: Dim doc As Word.Document
    Dim rs As ReportService
    Dim r As Word.Range: Dim tbl As Table
    
    Dim i As Integer: Dim j As Integer
    Dim resultFlag As Boolean: resultFlag = False     '��ʼ�������True��ʾ�ɹ�
   
    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    FileCopy ThisWorkbook.Path & "\" & templateFolderName & "\" & templateFileName, ThisWorkbook.Path & "\" & resultFolderName & "\" & resultFileName
    Set doc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & resultFolderName & "\" & resultFileName)
    '���"ͼ"��"��"
    doc.Application.CaptionLabels.Add Name:="ͼ": doc.Application.CaptionLabels.Add Name:="��"
       
    Dim pictureOffset As Integer: pictureOffset = 4   '��1����4��ͼƬ
    Dim widthArray(1 To tableColumnCounts) As Double    '�����������ܱ����
    widthArray(1) = 20: widthArray(2) = 30: widthArray(3) = 60: widthArray(4) = 80: widthArray(5) = 245: widthArray(6) = 60
    
    Set rs = New ReportService
    rs.SetReportBasicFormat doc
     '�²��ṹ
    rs.InsertResultSummaryAndDetailPictureTable sheetName:=InspectionSheetName, doc:=doc, InspectionResultTable:=SubStructureInspectionResultTable _
    , pictureOffset:=pictureOffset, InspectionResultTableName:="SubStructInspectionResultTable", InspectionResultPictureTableName:="SubStructInspectionResultPictureTable" _
    , StorePictures:=SubStructurePictures, PicturesFrontToEndIndex:=SubStructurePicturesFrontToEndIndex, DescriptionCounts:=SubStructureDescriptionCounts, widthArray:=widthArray, CheckBoxName:="SubStructureCheckBox"
    'Debug.Print SubStructureInspectionResultTable(1, 1)
    '�ϲ��ṹ
    rs.InsertResultSummaryAndDetailPictureTable sheetName:=InspectionSheetName, doc:=doc, InspectionResultTable:=SuperStructureInspectionResultTable _
    , pictureOffset:=pictureOffset, InspectionResultTableName:="SuperStructInspectionResultTable", InspectionResultPictureTableName:="SuperStructInspectionResultPictureTable" _
    , StorePictures:=SuperStructurePictures, PicturesFrontToEndIndex:=SuperStructurePicturesFrontToEndIndex, DescriptionCounts:=SuperStructureDescriptionCounts, widthArray:=widthArray, CheckBoxName:="SuperStructureCheckBox"
    'Debug.Print SubStructureInspectionResultTable(4, 2)
    '����ϵ
    Dim tableName As String
    Dim tableDataStartRow As Integer
    
    tableName = "BridgeDeckInspectionResultPictureTable"
    'NumRows���㷽��������ͼƬ��Ҫ����������������*2
    Set tbl = doc.Tables.Add(doc.Bookmarks(tableName).Range, numRows:=Fix((BridgeDeckPictureCounts + 1) / 2) * 2, numColumns:=2)  'NumRows+x��ʾ��ͷ
    
    '��������о�Ϊ�������оࡱ�������о�̫С�������ͼƬ��ʾ��ȫ
    tbl.Select
    With doc.Application.Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
    End With
    Dim insertedCounts As Integer: insertedCounts = 0 '�Ѳ���ͼƬ����
    '���β���ͼƬ
    For i = 1 To BridgeDeckDescriptionCounts
        If BridgeDeckInspectionResultTable(i, PictureCountsIndex) <> 0 Then    '�����ͼƬ�������
            For j = 1 To BridgeDeckInspectionResultTable(i, PictureCountsIndex)
                currPictureName = Dir(ThisWorkbook.Path & "\" & pictureFolderName & "\" & "*" & BridgeDeckPictures(i, j) & ".jpg")
                insertedCounts = insertedCounts + 1
                tbl.Cell(rs.GetTableRow(insertedCounts), rs.GetTableCol(insertedCounts)).Range.InlineShapes.AddPicture fileName:=ThisWorkbook.Path & "\" & pictureFolderName & "\" & currPictureName, LinkToFile:=False, SaveWithDocument:=True
                
                Set r = tbl.Cell(rs.GetTableRow(insertedCounts) + 1, rs.GetTableCol(insertedCounts)).Range
                r.Select
                '������ע
                With doc.Application.Selection
                    .MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
                    .TypeText "ͼ "
                    .Fields.Add doc.Application.Selection.Range, wdFieldEmpty, "STYLEREF 1 \s", False
                    .TypeText "-"
                    .Fields.Add doc.Application.Selection.Range, wdFieldEmpty, "SEQ ͼ \* ARABIC \s 1", False
                    If BridgeDeckInspectionResultTable(i, PictureCountsIndex) <> 1 Then    '�����ֻһ����Ƭ����Ӻ�׺
                        .TypeText " " & CStr(BridgeDeckInspectionResultTable(i, PictureDescriptionIndex) & "-" & CStr(j))
                    Else
                        .TypeText " " & CStr(BridgeDeckInspectionResultTable(i, PictureDescriptionIndex))
                    End If
                End With
            Next
        End If
    Next

    rs.SetTableBorderNoneLineStyle tbl
    tbl.Rows.Alignment = wdAlignRowCenter
    
    'û�в����������г����
    tableName = "BridgeDeckInspectionResultTable"
    If ActiveSheet.CheckBoxes("BridgeDeckCheckBox") = xlOff Then
        '���������ܱ��

        Set tbl = doc.Tables.Add(doc.Bookmarks(tableName).Range, numRows:=BridgeDeckDescriptionCounts + 1, numColumns:=tableColumnCounts) 'NumRows+x��ʾ��ͷ
        rs.SetTableBasicFormat tbl
        rs.SetTableColumnWidth tbl, BridgeDeckDescriptionCounts + 1, tableColumnCounts, widthArray
        
        'TODO����ȡ���º���
        'SetDispTableWidth tbl, DispUbound(i) + tableDataStartRow - 1    '���ø��п��
        tbl.Cell(1, 1).Range.InsertAfter "���": tbl.Cell(1, 2).Range.InsertAfter "λ��"
        tbl.Cell(1, 3).Range.InsertAfter "��������": tbl.Cell(1, 4).Range.InsertAfter "ȱ������"
        tbl.Cell(1, 5).Range.InsertAfter "��������": tbl.Cell(1, 6).Range.InsertAfter "ͼʾ���"
        
        For i = 1 To tableColumnCounts
            tbl.Cell(1, i).Range.Font.Bold = True
        Next
    
        tableDataStartRow = 2
        For i = 1 To BridgeDeckDescriptionCounts
            tbl.Cell(tableDataStartRow + i - 1, 1).Range.InsertAfter CStr(i)    '���
            tbl.Cell(tableDataStartRow + i - 1, 2).Range.InsertAfter BridgeDeckInspectionResultTable(i, PositionIndex)
            tbl.Cell(tableDataStartRow + i - 1, 3).Range.InsertAfter BridgeDeckInspectionResultTable(i, ComponentTypeIndex)
            tbl.Cell(tableDataStartRow + i - 1, 4).Range.InsertAfter BridgeDeckInspectionResultTable(i, DamageTypeIndex)
            
            tbl.Cell(tableDataStartRow + i - 1, 5).Range.InsertAfter rs.SetSquareMeterSuperscript(BridgeDeckInspectionResultTable(i, DamageDescriptionIndex))
            tbl.Cell(tableDataStartRow + i - 1, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            
            If BridgeDeckInspectionResultTable(i, PictureCountsIndex) <> 0 Then    '������Ϊ0
                 tbl.Cell(tableDataStartRow + i - 1, 6).Range.Select
                 If BridgeDeckInspectionResultTable(i, PictureCountsIndex) = 1 Then    'ֻ��1��
                      doc.Application.Selection.InsertCrossReference ReferenceType:="ͼ", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=BridgeDeckPicturesFrontToEndIndex(i, 1) + pictureOffset, _
                    InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
                ElseIf BridgeDeckInspectionResultTable(i, PictureCountsIndex) = 2 Then    '2��
                       doc.Application.Selection.InsertCrossReference ReferenceType:="ͼ", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=BridgeDeckPicturesFrontToEndIndex(i, 1) + pictureOffset, _
                    InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
                        doc.Application.Selection.TypeText vbCrLf
                       doc.Application.Selection.InsertCrossReference ReferenceType:="ͼ", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=BridgeDeckPicturesFrontToEndIndex(i, 2) + pictureOffset, _
                    InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
                Else
                       doc.Application.Selection.InsertCrossReference ReferenceType:="ͼ", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=BridgeDeckPicturesFrontToEndIndex(i, 1) + pictureOffset, _
                    InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
                    doc.Application.Selection.TypeText vbCrLf
                    doc.Application.Selection.TypeText "��"
                    doc.Application.Selection.TypeText vbCrLf
                       doc.Application.Selection.InsertCrossReference ReferenceType:="ͼ", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=BridgeDeckPicturesFrontToEndIndex(i, 2) + pictureOffset, _
                    InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
                 End If
            Else
                tbl.Cell(tableDataStartRow + i - 1, 6).Range.InsertAfter "/"
            End If

        Next
        rs.MergeSameColumnWithRefCol tbl:=tbl, col:=3, startRow:=2
        rs.MergeSameColumn tbl:=tbl, col:=2, startRow:=2
        rs.SetTableBorder tbl
        tbl.Rows.Alignment = wdAlignRowCenter
    Else
        doc.Application.Selection.GoTo what:=wdGoToBookmark, Name:=tableName
        doc.Application.Selection.Delete Unit:=wdCharacter, Count:=1
    End If

    doc.Save
    
    resultFlag = True
CloseWord:
    wordApp.Documents.Close
    wordApp.Quit
    
    Set wordApp = Nothing: Set doc = Nothing
    Set tbl = Nothing: Set rs = Nothing
    Set r = Nothing
    If resultFlag = True Then
        Debug.Print "���浼���ɹ���"
    Else
        Debug.Print "���浼��ʧ�ܣ�"
    End If
    
End Sub


Public Sub aa()
    Dim t As DropDown
    Set t = ActiveSheet.DropDowns("DropDown8")
    t.Delete
    t.AddItem "i1", 1
    t.AddItem "i2", 2

    Debug.Print t
End Sub
