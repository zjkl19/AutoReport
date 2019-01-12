Attribute VB_Name = "AutoDispStrainShare"
'删除当前sheet中所有ChartObject
Sub DeleteAllCharts()
    Dim c As ChartObject
    For Each c In ActiveSheet.ChartObjects

        c.Delete
    Next
End Sub
