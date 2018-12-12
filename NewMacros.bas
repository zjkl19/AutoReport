Attribute VB_Name = "NewMacros"
Sub Macro1()
Attribute Macro1.VB_Description = "宏由 林迪南 录制，时间: 2018/12/12"
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " 14"
'
' Macro1 Macro
' 宏由 林迪南 录制，时间: 2018/12/12
'

    ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmooth).Select
    ActiveChart.SetSourceData Source:=Range(Cells(1, 1), Cells(11, 3))
'    ActiveSheet.ChartObject.Chart.ChartStyle = 240
'
'    Application.Left = 8
'    Application.Top = 18
'    Application.WindowState = xlMaximized
'    Application.Left = 0
'    Application.Top = 0
'    Application.Width = 1920
'    Application.Height = 1040
End Sub
