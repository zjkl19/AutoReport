Attribute VB_Name = "AutoDispStrainShare"
'ɾ����ǰsheet������ChartObject
Sub DeleteAllCharts()
    Dim c As ChartObject
    For Each c In ActiveSheet.ChartObjects

        c.Delete
    Next
End Sub
