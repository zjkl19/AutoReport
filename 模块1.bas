Attribute VB_Name = "ģ��1"
Sub t1()

    Sheets("�ݸ�").Shapes.AddChart2(240, xlXYScatterSmooth).Select
    ActiveChart.SetSourceData Source:=Range("�ݸ�!$D$1:$F$11")
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)    '��������
    
    ActiveChart.ChartTitle.Delete
    
    ActiveChart.Axes(xlCategory).HasTitle = True
    ActiveChart.Axes(xlCategory).AxisTitle.Text = "����"
    ActiveChart.Axes(xlValue).HasTitle = True
    ActiveChart.Axes(xlValue).AxisTitle.Text = "�Ӷ�ֵ��mm��"
End Sub

'excel ������word��ǩ��ʾ��
Sub t2()

    Sheets("�ݸ�").Shapes.AddChart2(240, xlXYScatterSmooth).Select
    ActiveChart.SetSourceData Source:=Range("�ݸ�!$D$1:$F$11")
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)    '��������
    ActiveChart.Axes(xlCategory).HasTitle = True
    ActiveChart.Axes(xlCategory).AxisTitle.Text = "����"
    ActiveChart.Axes(xlValue).HasTitle = True
    ActiveChart.Axes(xlValue).AxisTitle.Text = "�Ӷ�ֵ��mm��"
    ActiveChart.CopyPicture
    

Dim wordApp As Word.Application
    Dim doc As Word.Document
    Dim r As Word.Range
    
    Set wordApp = CreateObject("Word.Application")
    wordApp.Documents.Open ThisWorkbook.Path & "\AutoReport1.docx"

    wordApp.Visible = False

    wordApp.ActiveDocument.Bookmarks("CH1").Range.Select
    wordApp.Selection.Paste
    wordApp.Documents.Save
    'wordApp.activedocument.SaveAs2 ThisWorkbook.Path & "\AutoReport1.docx"
    wordApp.Documents.Close
    wordApp.Quit
    
    Set wordApp = Nothing
    

End Sub
'Word DOCVARIABLE replacement
Sub t3()



    Dim wordApp As Word.Application
        Dim doc As Word.Document
        Dim r As Word.Range
        
        Set wordApp = CreateObject("Word.Application")
        wordApp.Documents.Open ThisWorkbook.Path & "\AutoReportTemplate.docx"
    
        wordApp.Visible = False
    
        
        wordApp.ActiveDocument.Bookmarks("dispSummary1").Range.InsertAfter "111"
        

        '���Բ��������
        'wordApp.ActiveDocument.Tables.Add wordApp.ActiveDocument.Bookmarks("dispSummary1").Range, NumRows:=14 + 1, NumColumns:=7
        
        'Debug.Print wordApp.ActiveDocument.Variables("tb1").value
        wordApp.ActiveDocument.Variables("tb1").value = "testTb1"
        wordApp.ActiveDocument.Fields.Update
        'wordApp.Selection.Paste    '�������������ճ����"Select"�Ĳ���
        
'        Dim bk
'        For Each bk In wordApp.ActiveDocument.Bookmarks
'            bk.Delete
'        Next
'
'        Dim oFld
'        For Each oFld In ActiveDocument.Fields
'            If oFld.Type = wdFieldDocVariable Then
'                 oFld.Update
'            End If
'        Next oFld
        wordApp.Documents.Save
        
        wordApp.ActiveDocument.SaveAs2 ThisWorkbook.Path & "\AutoReportResult.docx"
        wordApp.Documents.Close
        wordApp.Quit
        
        Set wordApp = Nothing
        
    
    End Sub

Sub ��1()
Attribute ��1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��1 ��
'

'
    'Range("A1:C11").Select
    Dim myChart
    myChart = ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmooth)
    myChart.SetSourceData Source:=Range("�ݸ�!$A$1:$C$11")
    ActiveChart.Legend.Select
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    ActiveChart.HasLegend = True
    ActiveChart.Legend.Select
    ActiveChart.Legend.IncludeInLayout = False
    ActiveChart.Legend.Select
    Selection.Position = xlRight
    ActiveChart.Legend.Select
    Selection.Left = 48
    Selection.Top = 46.107
    Selection.Left = 63
    Selection.Top = 44.107
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "����"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "����"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 3).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 3).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.Axes(xlValue).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "�Ӷ�ֵ��mm��"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "�Ӷ�ֵ��mm��"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 7).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 4).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(5, 2).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(7, 1).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
    ActiveChart.ChartTitle.Select
    Selection.Delete
End Sub
