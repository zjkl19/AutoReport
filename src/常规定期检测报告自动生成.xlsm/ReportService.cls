VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReportService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'打开报告
'fileName：含路径，如"\自动生成的常规定期检测报告\自动生成的常规定期检测报告.docx"
Public Sub OpenReport(ByVal fileName)
    Dim wordApp As Word.Application
    Set wordApp = CreateObject("Word.Application")
    wordApp.Documents.Open fileName:=fileName, ReadOnly:=False
    wordApp.Visible = True
    wordApp.Activate
End Sub

'设置报告基本格式
'建造者模式
Public Sub SetReportBasicFormat(ByRef doc As Word.Document)
    SetReportMainBodyFormat doc
    SetReportCaptionsFormat doc
End Sub

'设置报告"正文"样式格式
Public Sub SetReportMainBodyFormat(ByRef doc As Word.Document)
    With doc.Styles.Item("正文")
        With .Font
            .NameAscii = "Times New Roman"
            .NameOther = "Times New Roman"
            .NameFarEast = "楷体_GB2312"
            .Size = 12                           '小四
        End With
    End With
End Sub

'设置报告"题注"样式格式
Public Sub SetReportCaptionsFormat(ByRef doc As Word.Document)
    With doc.Styles.Item("题注")
        With .Font
            .NameAscii = "Times New Roman"
            .NameOther = "Times New Roman"
            .NameFarEast = "楷体_GB2312"
            .Size = 12                           '小四
        End With
    End With
End Sub

'初始化检测结果变量
'sheetName：标签名
'structureType：结构类型：桥面系，上部结构，下部结构
'InspectionResultTable()：2维检测结果数组
'StorePictures()：每张照片尾号
'PicturesFrontToEndIndex()：照片首尾索引
'DescriptionCounts：描述数量统计值
'PictureCounts：照片数量统计值
'SpanCounts：跨数统计值
Public Sub InitInspectionResult(ByVal sheetName As String, ByVal structureType As String, ByRef InspectionResultTable() As String _
    , ByRef StorePictures() As String, ByRef PicturesFrontToEndIndex() As Integer _
    , ByRef DescriptionCounts As Integer, ByRef PictureCounts As Integer, ByRef SpanCounts As Integer)
    Dim startRow As Integer: startRow = 2
    Dim currRow As Integer: currRow = startRow
    Dim i As Integer: Dim j As Integer: i = 1

    DescriptionCounts = 0: PictureCounts = 0: SpanCounts = 1
    Dim pictureNoOut As String: Dim singlePictureCounts As Integer
    
    Dim pictures
    '读入桥面系信息
    While Sheets(sheetName).Cells(currRow, 11) <> ""    '只要不为空就继续读取下一行
        If Sheets(sheetName).Cells(currRow, 11) = structureType Then    '符合条件则存储
            InspectionResultTable(i, 1) = Sheets(sheetName).Cells(currRow, BridgePartColumn): InspectionResultTable(i, 2) = Sheets(sheetName).Cells(currRow, PositionColumn)
            InspectionResultTable(i, 3) = Sheets(sheetName).Cells(currRow, ComponentTypeColumn): InspectionResultTable(i, 4) = Sheets(sheetName).Cells(currRow, DamageTypeColumn)
            InspectionResultTable(i, 5) = Sheets(sheetName).Cells(currRow, DamageDescriptionColumn): InspectionResultTable(i, 6) = Sheets(sheetName).Cells(currRow, PictureDescriptionColumn)
            InspectionResultTable(i, 7) = CStr(Sheets(sheetName).Cells(currRow, PictureNoColumn))
            singlePictureCounts = TranslatePictureNo(InspectionResultTable(i, 7), pictureNoOut)
            InspectionResultTable(i, 8) = CStr(pictureNoOut)
            InspectionResultTable(i, 9) = UBound(Split(CStr(pictureNoOut), ",")) + 1    '相片数量
            
            'TODO：抽取计算跨数的方法
            '桥跨默认只有1行，若从第2行起，改行桥跨名与上一行不同，则+1
            If i > 1 Then
                If InspectionResultTable(i, 2) <> InspectionResultTable(i - 1, 2) Then
                    SpanCounts = SpanCounts + 1    'Debug.Print SpanCounts
                End If
            End If
            'Debug.Print InspectionResultTable(i, 9)
            
            If pictureNoOut <> "" Then  '如果图片数量不为0
                pictures = Split(CStr(pictureNoOut), ",")
                    For j = 1 To CInt(InspectionResultTable(i, 9))    '依次存储各个相片尾号
                        StorePictures(i, j) = CStr(pictures(j - 1))
                    Next j
            End If
            
            If pictureNoOut = "" Then  '图片数量为0
                PicturesFrontToEndIndex(i, 1) = 0: PicturesFrontToEndIndex(i, 2) = 0
            Else
                If i = 1 Then  '如果是第1张
                    PicturesFrontToEndIndex(i, 1) = 1: PicturesFrontToEndIndex(i, 2) = 1 + CInt(InspectionResultTable(i, 9)) - 1
                Else
                    PicturesFrontToEndIndex(i, 1) = SumArray(InspectionResultTable, 9, i - 1) + 1
                    PicturesFrontToEndIndex(i, 2) = SumArray(InspectionResultTable, 9, i - 1) + CInt(InspectionResultTable(i, 9))
                End If
            End If
            'Debug.Print PicturesFrontToEndIndex(i, 1)
            'Debug.Print PicturesFrontToEndIndex(i, 2)
    
            PictureCounts = PictureCounts + singlePictureCounts     '算法：UBound(数组名)-LBound(数组名)+1
            
            i = i + 1: DescriptionCounts = DescriptionCounts + 1
        End If
        currRow = currRow + 1
    Wend

    'Debug.Print BridgeDeckPictureCounts
End Sub

'插入汇总表格和详细图片表格
'sheetName：标签名
'doc：已打开的word文档
'InspectionResultTable()：2维检测结果数组
'pictureOffset：图片偏移量
'InspectionResultTableName：检测结果表格名称
'InspectionResultPictureTableName：检测结果图片表名称
'StorePictures：存储的照片
'PicturesFrontToEndIndex()：照片首尾索引
'DescriptionCounts：描述数量
'widthArray：设置表格列宽度的数组
'CheckBoxName：决定是否插入汇总表格的CheckBox名称
Public Sub InsertResultSummaryAndDetailPictureTable(ByVal sheetName As String, ByRef doc As Word.Document, ByRef InspectionResultTable() As String _
    , ByVal pictureOffset As Integer, ByVal InspectionResultTableName As String, ByVal InspectionResultPictureTableName As String _
    , ByRef StorePictures() As String, ByRef PicturesFrontToEndIndex() As Integer, ByVal DescriptionCounts As Integer, ByRef widthArray() As Double, ByVal CheckBoxName As String)
    
    Const tableColumnCounts As Integer = 6  '表格列数（未合并任意单元格）
    Dim currPictureName As String    '当前照片名称
    
    Dim pictureFolderName As String: pictureFolderName = "常规定期检测照片"
    
    Dim r As Word.Range: Dim tbl As Table
    Dim i As Integer: Dim j As Integer
    Dim tableName As String: Dim tableDataStartRow As Integer
    
    tableName = InspectionResultPictureTableName    '举例："BridgeDeckInspectionResultPictureTable"
    'NumRows计算方法：由于图片还要包含描述，故行数*2
    Set tbl = doc.Tables.Add(doc.Bookmarks(tableName).Range, numRows:=Fix(((DescriptionCounts + 1) / 2)) * 2, numColumns:=2)  'NumRows+x表示表头
    
    tbl.Select
    With doc.Application.Selection.ParagraphFormat    '调整表格行距为“单倍行距”，否则行距太小的情况下图片显示不全
        .LineSpacingRule = wdLineSpaceSingle
    End With
    Dim insertedCounts As Integer: insertedCounts = 0 '已插入图片数量
    '依次插入图片
    For i = 1 To DescriptionCounts
        If InspectionResultTable(i, PictureCountsIndex) <> 0 Then    '如果有图片，则插入
            For j = 1 To InspectionResultTable(i, PictureCountsIndex)
                currPictureName = Dir(ThisWorkbook.Path & "\" & pictureFolderName & "\" & "*" & StorePictures(i, j) & ".jpg")
                insertedCounts = insertedCounts + 1
                tbl.Cell(GetTableRow(insertedCounts), GetTableCol(insertedCounts)).Range.InlineShapes.AddPicture fileName:=ThisWorkbook.Path & "\" & pictureFolderName & "\" & currPictureName, LinkToFile:=False, SaveWithDocument:=True
                
                Set r = tbl.Cell(GetTableRow(insertedCounts) + 1, GetTableCol(insertedCounts)).Range
                r.Select
                '插入题注
                With doc.Application.Selection
                    .MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdMove
                    .TypeText "图 "
                    .Fields.Add doc.Application.Selection.Range, wdFieldEmpty, "STYLEREF 1 \s", False
                    .TypeText "-"
                    .Fields.Add doc.Application.Selection.Range, wdFieldEmpty, "SEQ 图 \* ARABIC \s 1", False
                    If InspectionResultTable(i, PictureCountsIndex) <> 1 Then    '如果不只一张照片，则加后缀
                        .TypeText " " & CStr(InspectionResultTable(i, PictureDescriptionIndex) & "-" & CStr(j))
                    Else
                        .TypeText " " & CStr(InspectionResultTable(i, PictureDescriptionIndex))
                    End If
                End With
            Next
        End If
    Next

    SetTableBorderNoneLineStyle tbl
    tbl.Rows.Alignment = wdAlignRowCenter
    
    '没有病害则无需列出表格
    tableName = InspectionResultTableName    '举例："BridgeDeckInspectionResultTable"
    If Sheets(sheetName).CheckBoxes(CheckBoxName) = xlOff Then
        '插入结果汇总表格

        Set tbl = doc.Tables.Add(doc.Bookmarks(tableName).Range, numRows:=DescriptionCounts + 1, numColumns:=tableColumnCounts) 'NumRows+x表示表头
        SetTableBasicFormat tbl
        SetTableColumnWidth tbl, DescriptionCounts + 1, tableColumnCounts, widthArray
        
        'TODO：抽取以下函数
        'SetDispTableWidth tbl, DispUbound(i) + tableDataStartRow - 1    '设置各列宽度
        tbl.Cell(1, 1).Range.InsertAfter "序号": tbl.Cell(1, 2).Range.InsertAfter "位置"
        tbl.Cell(1, 3).Range.InsertAfter "构件类型": tbl.Cell(1, 4).Range.InsertAfter "缺损类型"
        tbl.Cell(1, 5).Range.InsertAfter "病害描述": tbl.Cell(1, 6).Range.InsertAfter "图示编号"
        
        For i = 1 To tableColumnCounts
            tbl.Cell(1, i).Range.Font.Bold = True
        Next
    
        tableDataStartRow = 2
        For i = 1 To DescriptionCounts
            tbl.Cell(tableDataStartRow + i - 1, 1).Range.InsertAfter CStr(i)    '序号
            tbl.Cell(tableDataStartRow + i - 1, 2).Range.InsertAfter InspectionResultTable(i, PositionIndex)
            tbl.Cell(tableDataStartRow + i - 1, 3).Range.InsertAfter InspectionResultTable(i, ComponentTypeIndex)
            tbl.Cell(tableDataStartRow + i - 1, 4).Range.InsertAfter InspectionResultTable(i, DamageTypeIndex)
            
            tbl.Cell(tableDataStartRow + i - 1, 5).Range.InsertAfter SetSquareMeterSuperscript(InspectionResultTable(i, DamageDescriptionIndex))
            tbl.Cell(tableDataStartRow + i - 1, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            
            If InspectionResultTable(i, PictureCountsIndex) <> 0 Then    '数量不为0
                    tbl.Cell(tableDataStartRow + i - 1, 6).Range.Select
                 If InspectionResultTable(i, PictureCountsIndex) = 1 Then    '只有1张
                      doc.Application.Selection.InsertCrossReference ReferenceType:="图", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=PicturesFrontToEndIndex(i, 1) + pictureOffset, _
                    InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
                ElseIf InspectionResultTable(i, PictureCountsIndex) = 2 Then    '2张
                       doc.Application.Selection.InsertCrossReference ReferenceType:="图", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=PicturesFrontToEndIndex(i, 1) + pictureOffset, _
                    InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
                        doc.Application.Selection.TypeText vbCrLf
                       doc.Application.Selection.InsertCrossReference ReferenceType:="图", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=PicturesFrontToEndIndex(i, 2) + pictureOffset, _
                    InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
                Else
                       doc.Application.Selection.InsertCrossReference ReferenceType:="图", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=PicturesFrontToEndIndex(i, 1) + pictureOffset, _
                    InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
                       doc.Application.Selection.TypeText vbCrLf
                       doc.Application.Selection.TypeText "～"
                       doc.Application.Selection.TypeText vbCrLf
                       doc.Application.Selection.InsertCrossReference ReferenceType:="图", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=PicturesFrontToEndIndex(i, 2) + pictureOffset, _
                    InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
                 End If
            Else
                tbl.Cell(tableDataStartRow + i - 1, 6).Range.InsertAfter "/"
            End If

        Next
        MergeSameColumnWithRefCol tbl:=tbl, col:=3, startRow:=2
        MergeSameColumn tbl:=tbl, col:=2, startRow:=2
        SetTableBorder tbl
        tbl.Rows.Alignment = wdAlignRowCenter
    Else
        doc.Application.Selection.GoTo what:=wdGoToBookmark, Name:=tableName
        doc.Application.Selection.Delete Unit:=wdCharacter, Count:=1
    End If
End Sub

'统计数组指定维数（至指定上界）的和
'arr():准备统计的2维数组（下届为1）
'countDimension:统计的维度
'countUbound:计算的上界
'举例：
'BridgeDeckInspectionResultTable(1,9)=2:BridgeDeckInspectionResultTable(2,9)=3 :BridgeDeckInspectionResultTable(3,9)=1
'桥面系第1个描述2张照片，第2个描述3张照片，第3个描述1张照片
'SumArray(BridgeDeckInspectionResultTable,9,2)==5
Public Function SumArray(ByRef arr() As String, ByVal countDimension As Integer, ByVal countUbound As Integer) As Integer
    Dim i As Integer: Dim sum As Integer: sum = 0
    For i = 1 To countUbound
        sum = sum + CInt(arr(i, countDimension))
    Next
    SumArray = sum
End Function

'转换原始输入的照片编号数据
'算法：
'返回：照片数量
'示例：
'示例输入：58,59,46-48,62-67,56
'示例输出：58,59,46,47,48,62,63,64,65,66,67,56
'TODO：检查输入是否正确
'Public 改成 private
Public Function TranslatePictureNo(ByVal pictureNoIn As String, ByRef pictureNoOut As String) As Integer
    If InStr(pictureNoIn, "-") > 0 Then    '存在"-"则必须转换
        'TODO：修改如下代码
        pictureNoOut = pictureNoIn
    Else
        pictureNoOut = pictureNoIn
    End If
    TranslatePictureNo = UBound(Split(CStr(pictureNoIn), ",")) + 1
'        Dim TargetRange As Object    'Dim TargetRange As range 类型和doc.Application.Selection.Range不匹配
'
'    Dim mRegExp As Object       '正则表达式对象
'    Dim mMatches As Object      '匹配字符串集合对象
'    Dim mMatch As Object        '匹配字符串
'
'    Dim matchResult As String
'
'    doc.Application.Selection.WholeStory    '全选
'    Set TargetRange = doc.Application.Selection.Range
'
'    Set mRegExp = CreateObject("Vbscript.Regexp")    'New RegExp
'    With mRegExp
'        .Global = True                              'True表示匹配所有, False表示仅匹配第一个符合项
'        .IgnoreCase = True                          'True表示不区分大小写, False表示区分大小写
'        '(?<=工程概况)[\S\s]{2,1000}(，|。)(?=[\s\S]{1,20}图 1)
'        '概况\n[\S\s]+[,，。]\n?[^\n]+图\s?1\S+。
'        '工程概况\r[\S\s]{2,2000}桥梁结构布置
'        '.Pattern = "概况\r[\S\s]+[,，。]\r?[^\r]+图\s?1\S*。"
'        .Pattern = "概况\r[\S\s]+?[,，。]\r?[^\r]+?图\s?1\S*。"   '匹配字符模式  (工程概况)+([\s\S]*)+(桥梁结构布置) (?<=项目名称：).*?\r|(?<=工程名称：).*?\r
'        Set mMatches = .Execute(TargetRange.Text)   '执行正则查找，返回所有匹配结果的集合，若未找到，则为空
'        For Each mMatch In mMatches
'            matchResult = CStr(mMatch.Value)
'        Next
'    End With
'
'    With mRegExp
'        .Global = True
'        .IgnoreCase = True
'        .Pattern = "[,，。]\r?[^\r,，。]+图\s?1\S*。"
'        matchResult = .Replace(matchResult, "。")
'
'    End With
'
'    Set mRegExp = Nothing
'    Set mMatches = Nothing
'    Set TargetRange = Nothing
'    'Debug.Print matchResult
'    matchResult = Replace(matchResult, "概况", "")
'    matchResult = Replace(matchResult, vbCr, "")
'    matchResult = Replace(matchResult, vbLf, "")
'    GetProjectOverviewFromReport = matchResult
'
End Function

'判断对应病害是否有照片
'算法：有内容（图片编号）则为有照片，否则为无照片
Public Function HasPicture(ByVal pictureNo As String) As Boolean
    If pictureNo = "" Then
        HasPicture = False
    Else
        HasPicture = True
    End If
End Function

'计算第n张图片对应表格的行号（考虑图名占1行）
Public Function GetTableRow(ByVal n As Integer) As Integer
    If n Mod 2 = 0 Then
        GetTableRow = n - 1
    Else
        GetTableRow = n
    End If
End Function

'计算第n张图片对应表格的列号（考虑图名占1行）
Public Function GetTableCol(ByVal n As Integer) As Integer
    If n Mod 2 = 0 Then
        GetTableCol = 2
    Else
        GetTableCol = 1
    End If
End Function

'将字符串中"m2"（平方米）替换为"m2（上标）"
'根据https://en.wikipedia.org/wiki/Latin_script_in_Unicode，ChrW(&H00B2)为2的上标
Public Function SetSquareMeterSuperscript(ByVal stringIn)
    SetSquareMeterSuperscript = Replace(stringIn, "m2", "m" & ChrW(&HB2))
End Function

'合并列中相同内容的单元格
'要求：原始表格必须是没有合并过的
'算法：合并时，除了本列单元格（构件类型）一致外，前一列（位置）也必须一致
'col：需要合并的列
'startRow：开始合并的行，默认从第2行开始
Public Sub MergeSameColumnWithRefCol(ByRef tbl As Word.Table, ByVal col As Integer, Optional startRow As Integer = 2)
    On Error Resume Next
    Dim refCol As Integer    '参考列，默认为前一列（位置列）
    refCol = col - 1
    Dim totalRows As Integer: totalRows = tbl.Rows.Count
    Dim mergeLength As Integer    '合并长度
    mergeLength = 0
    Dim currRow As Integer    '当前行
    Dim currContent As String    '当前单元格的内容
    Dim currRefContent As String    '当前参考列的内容
    Dim i As Integer: Dim j As Integer
    
    '逻辑较复杂，需作图理解算法
    currRow = startRow
    For i = startRow To totalRows - 1
        mergeLength = 0
        currContent = tbl.Cell(currRow, col).Range.Text: currRefContent = tbl.Cell(currRow, refCol).Range.Text
        
        '计算合并的长度
        For j = currRow + 1 To totalRows
            If currContent = tbl.Cell(j, col).Range.Text And currRefContent = tbl.Cell(j, refCol).Range.Text Then
                mergeLength = mergeLength + 1
            Else
                Exit For
            End If
        Next j
        
        If mergeLength <> 0 Then
            tbl.Cell(currRow, col).Merge tbl.Cell(currRow + mergeLength, col)
            tbl.Cell(currRow, col).Range.Text = Mid(currContent, 1, Len(currContent) - 2)   '最后一个回车换行符去掉（CrLf占两个位置）
            currRow = currRow + mergeLength
        End If
        
        currRow = currRow + 1    '无论是否合并，都+1
        
        If currRow >= totalRows Then
            Exit For
        End If
    Next i

End Sub

'合并列中相同内容的单元格
'要求：原始表格必须是没有合并过的
'col：需要合并的列
'startRow：开始合并的行，默认从第2行开始
Public Sub MergeSameColumn(ByRef tbl As Word.Table, ByVal col As Integer, Optional startRow As Integer = 2)
    On Error Resume Next
    Dim totalRows As Integer: totalRows = tbl.Rows.Count
    Dim mergeLength As Integer    '合并长度
    mergeLength = 0
    Dim currRow As Integer    '当前行
    Dim currContent As String    '当前单元格的内容
    Dim i As Integer: Dim j As Integer
    
    '逻辑较复杂，需作图理解算法
    currRow = startRow
    For i = startRow To totalRows - 1
        mergeLength = 0
        currContent = tbl.Cell(currRow, col).Range.Text
        
        '计算合并的长度
        For j = currRow + 1 To totalRows
            If currContent = tbl.Cell(j, col).Range.Text Then
                mergeLength = mergeLength + 1
            Else
                Exit For
            End If
        Next j
        
        If mergeLength <> 0 Then
            tbl.Cell(currRow, col).Merge tbl.Cell(currRow + mergeLength, col)
            tbl.Cell(currRow, col).Range.Text = Mid(currContent, 1, Len(currContent) - 2)   '最后一个回车换行符去掉（CrLf占两个位置）
            currRow = currRow + mergeLength
        End If
        
        currRow = currRow + 1    '无论是否合并，都+1
        
        If currRow >= totalRows Then
            Exit For
        End If
    Next i

End Sub

'设置图片交叉引用
'pictureOffset：图片偏移量
Public Sub SetPictureCrossReferences(ByRef wordApp As Word.Application, Optional pictureOffset As Integer = 0)
'    Dim i As Integer
'    For i = 1 To NWCs
'        With wordApp.Application
'            '插入交叉引用
'
'            .Selection.Find.Execute Replace:=wdReplaceNone
'            .Selection.TypeBackspace
'            .Selection.InsertCrossReference ReferenceType:="表", ReferenceKind:=wdOnlyLabelAndNumber, ReferenceItem:=rs.GetDispCrossReferenceItem(i, GlobalWC, StrainGlobalWC, strainNWCs, tableOffset), _
'        InsertAsHyperlink:=True, IncludePosition:=False, SeparateNumbers:=False, SeparatorString:=""
'        End With
'    Next
End Sub

'设置table基本格式（单元格未合并前）
Public Sub SetTableBasicFormat(ByRef tbl As Table)
    On Error Resume Next
    tbl.Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
End Sub

'设置table各列宽度
Public Sub SetTableColumnWidth(ByRef tbl As Table, ByVal numRows As Integer, ByVal numColumns As Integer, ByRef widthArray() As Double)
    On Error Resume Next
    Dim i As Integer: Dim j As Integer
    tbl.AllowAutoFit = False
    For i = 1 To numColumns
        For j = 1 To numRows
             tbl.Cell(j, i).SetWidth ColumnWidth:=widthArray(i), RulerStyle:=wdAdjustFirstColumn
        Next j
    Next i

End Sub

'设置table的边界线
Public Sub SetTableBorder(ByRef tbl As Word.Table)
    On Error Resume Next
    With tbl
        With .Borders
            .InsideLineStyle = wdLineStyleSingle
            .OutsideLineStyle = wdLineStyleSingle
        End With
    End With
    
    With tbl
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth150pt
        End With
    End With
End Sub

'设置table没有边界线
Public Sub SetTableBorderNoneLineStyle(ByRef tbl As Word.Table)
    On Error Resume Next
    With tbl
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    End With
End Sub
