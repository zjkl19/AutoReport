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
    fileName = ThisWorkbook.Path & "\自动生成的常规定期检测报告\自动生成的桥梁常规定期检测报告.docx"
    rs.OpenReport fileName
    Set rs = Nothing
End Sub

'请先计算，再生成Word报告
Public Sub AutoReport()
    
    GetInspectionResult    '获取检测结果
    
    Const tableColumnCounts As Integer = 6  '表格列数（未合并任意单元格）
    
    Dim templateFolderName As String: Dim templateFileName As String
    templateFolderName = "常规定期检测报告模板": templateFileName = "桥梁常规定期检测报告模板.docx"
    Dim resultFolderName As String: Dim resultFileName As String
    resultFolderName = "自动生成的常规定期检测报告": resultFileName = "自动生成的桥梁常规定期检测报告.docx"
    Dim pictureFolderName As String: pictureFolderName = "常规定期检测照片"
    Dim currPictureName As String    '当前照片名称
    
    Dim wordApp As Word.Application: Dim doc As Word.Document
    Dim rs As ReportService
    Dim r As Word.Range: Dim tbl As Table
    
    Dim i As Integer: Dim j As Integer
    Dim resultFlag As Boolean: resultFlag = False     '初始化结果，True表示成功
   
    On Error GoTo CloseWord:
    Set wordApp = CreateObject("Word.Application")
    FileCopy ThisWorkbook.Path & "\" & templateFolderName & "\" & templateFileName, ThisWorkbook.Path & "\" & resultFolderName & "\" & resultFileName
    Set doc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & resultFolderName & "\" & resultFileName)
    '添加"图"，"表"
    doc.Application.CaptionLabels.Add Name:="图": doc.Application.CaptionLabels.Add Name:="表"
       
    Dim pictureOffset As Integer: pictureOffset = 4   '第1章有4张图片
    Dim widthArray(1 To tableColumnCounts) As Double    '控制描述汇总表格宽度
    widthArray(1) = 20: widthArray(2) = 30: widthArray(3) = 60: widthArray(4) = 80: widthArray(5) = 245: widthArray(6) = 60
    
    Set rs = New ReportService
    rs.SetReportBasicFormat doc
     '下部结构
    rs.InsertResultSummaryAndDetailPictureTable sheetName:=InspectionSheetName, doc:=doc, InspectionResultTable:=SubStructureInspectionResultTable _
    , pictureOffset:=pictureOffset, InspectionResultTableName:="SubStructInspectionResultTable", InspectionResultPictureTableName:="SubStructInspectionResultPictureTable" _
    , StorePictures:=SubStructurePictures, PicturesFrontToEndIndex:=SubStructurePicturesFrontToEndIndex, DescriptionCounts:=SubStructureDescriptionCounts, widthArray:=widthArray, CheckBoxName:="SubStructureCheckBox"
    'Debug.Print SubStructureInspectionResultTable(1, 1)
    '上部结构
    rs.InsertResultSummaryAndDetailPictureTable sheetName:=InspectionSheetName, doc:=doc, InspectionResultTable:=SuperStructureInspectionResultTable _
    , pictureOffset:=pictureOffset, InspectionResultTableName:="SuperStructInspectionResultTable", InspectionResultPictureTableName:="SuperStructInspectionResultPictureTable" _
    , StorePictures:=SuperStructurePictures, PicturesFrontToEndIndex:=SuperStructurePicturesFrontToEndIndex, DescriptionCounts:=SuperStructureDescriptionCounts, widthArray:=widthArray, CheckBoxName:="SuperStructureCheckBox"
    'Debug.Print SubStructureInspectionResultTable(4, 2)
    '桥面系
    Dim tableName As String
    Dim tableDataStartRow As Integer
    
    tableName = "BridgeDeckInspectionResultPictureTable"
    'NumRows计算方法：由于图片还要包含描述，故行数*2
    Set tbl = doc.Tables.Add(doc.Bookmarks(tableName).Range, numRows:=Fix((BridgeDeckPictureCounts + 1) / 2) * 2, numColumns:=2)  'NumRows+x表示表头
    
    '调整表格行距为“单倍行距”，否则行距太小的情况下图片显示不全
    tbl.Select
    With doc.Application.Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
    End With
    Dim insertedCounts As Integer: insertedCounts = 0 '已插入图片数量
    '依次插入图片
    For i = 1 To BridgeDeckDescriptionCounts
        If BridgeDeckInspectionResultTable(i, PictureCountsIndex) <> 0 Then    '如果有图片，则插入
            For j = 1 To BridgeDeckInspectionResultTable(i, PictureCountsIndex)
                currPictureName = Dir(ThisWorkbook.Path & "\" & pictureFolderName & "\" & "*" & BridgeDeckPictures(i, j) & ".jpg")
                insertedCounts = insertedCounts + 1
                tbl.Cell(rs.GetTableRow(insertedCounts), rs.GetTableCol(insertedCounts)).Range.InlineShapes.AddPicture fileName:=ThisWorkbook.Path & "\" & pictureFolderName & "\" & currPictureName, LinkToFile:=False, SaveWithDocument:=True
                
                Set r = tbl.Cell(rs.GetTableRow(insertedCounts) + 1, rs.GetTableCol(insertedCounts)).Range
                r.Select
                '插入题注
                With doc.Application.Selection
                    .MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
                    .TypeText "图 "
                    .Fields.Add doc.Application.Selection.Range, wdFieldEmpty, "STYLEREF 1 \s", False
                    .TypeText "-"
                    .Fields.Add doc.Application.Selection.Range, wdFieldEmpty, "SEQ 图 \* ARABIC \s 1", False
                    If BridgeDeckInspectionResultTable(i, PictureCountsIndex) <> 1 Then    '如果不只一张照片，则加后缀
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
    
    '没有病害则无需列出表格
    tableName = "BridgeDeckInspectionResultTable"
    If ActiveSheet.CheckBoxes("BridgeDeckCheckBox") = xlOff Then
        '插入结果汇总表格

        Set tbl = doc.Tables.Add(doc.Bookmarks(tableName).Range, numRows:=BridgeDeckDescriptionCounts + 1, numColumns:=tableColumnCounts) 'NumRows+x表示表头
        rs.SetTableBasicFormat tbl
        rs.SetTableColumnWidth tbl, BridgeDeckDescriptionCounts + 1, tableColumnCounts, widthArray
        
        'TODO：抽取以下函数
        'SetDispTableWidth tbl, DispUbound(i) + tableDataStartRow - 1    '设置各列宽度
        tbl.Cell(1, 1).Range.InsertAfter "序号": tbl.Cell(1, 2).Range.InsertAfter "位置"
        tbl.Cell(1, 3).Range.InsertAfter "构件类型": tbl.Cell(1, 4).Range.InsertAfter "缺损类型"
        tbl.Cell(1, 5).Range.InsertAfter "病害描述": tbl.Cell(1, 6).Range.InsertAfter "图示编号"
        
        For i = 1 To tableColumnCounts
            tbl.Cell(1, i).Range.Font.Bold = True
        Next
    
        tableDataStartRow = 2
        For i = 1 To BridgeDeckDescriptionCounts
            tbl.Cell(tableDataStartRow + i - 1, 1).Range.InsertAfter CStr(i)    '序号
            tbl.Cell(tableDataStartRow + i - 1, 2).Range.InsertAfter BridgeDeckInspectionResultTable(i, PositionIndex)
            tbl.Cell(tableDataStartRow + i - 1, 3).Range.InsertAfter BridgeDeckInspectionResultTable(i, ComponentTypeIndex)
            tbl.Cell(tableDataStartRow + i - 1, 4).Range.InsertAfter BridgeDeckInspectionResultTable(i, DamageTypeIndex)
            
            tbl.Cell(tableDataStartRow + i - 1, 5).Range.InsertAfter rs.SetSquareMeterSuperscript(BridgeDeckInspectionResultTable(i, DamageDescriptionIndex))
            tbl.Cell(tableDataStartRow + i - 1, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            
            If BridgeDeckInspectionResultTable(i, PictureCountsIndex) <> 0 Then    '数量不为0
                 tbl.Cell(tableDataStartRow + i - 1, 6).Range.Select
                 If BridgeDeckInspectionResultTable(i, PictureCountsIndex) = 1 Then    '只有1张
                      doc.Application.Selection.InsertCrossReference ReferenceType:="图", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=BridgeDeckPicturesFrontToEndIndex(i, 1) + pictureOffset, _
                    InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
                ElseIf BridgeDeckInspectionResultTable(i, PictureCountsIndex) = 2 Then    '2张
                       doc.Application.Selection.InsertCrossReference ReferenceType:="图", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=BridgeDeckPicturesFrontToEndIndex(i, 1) + pictureOffset, _
                    InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
                        doc.Application.Selection.TypeText vbCrLf
                       doc.Application.Selection.InsertCrossReference ReferenceType:="图", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=BridgeDeckPicturesFrontToEndIndex(i, 2) + pictureOffset, _
                    InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
                Else
                       doc.Application.Selection.InsertCrossReference ReferenceType:="图", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=BridgeDeckPicturesFrontToEndIndex(i, 1) + pictureOffset, _
                    InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
                    doc.Application.Selection.TypeText vbCrLf
                    doc.Application.Selection.TypeText "～"
                    doc.Application.Selection.TypeText vbCrLf
                       doc.Application.Selection.InsertCrossReference ReferenceType:="图", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=BridgeDeckPicturesFrontToEndIndex(i, 2) + pictureOffset, _
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
        Debug.Print "报告导出成功！"
    Else
        Debug.Print "报告导出失败！"
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
