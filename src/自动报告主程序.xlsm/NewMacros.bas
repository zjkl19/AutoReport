Attribute VB_Name = "NewMacros"
Sub Macro1()
Attribute Macro1.VB_Description = "���� Administrator ¼�ƣ�ʱ��: 2019/03/23"
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " 14"
'
' Macro1 Macro
' ���� Administrator ¼�ƣ�ʱ��: 2019/03/23
'

'
    Range("M9:R29").Select
    With ActiveSheet.Sort
        With .SortFields
            .Clear
            .Add Key:=Range("Q9:Q29"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:=""
        End With
        .Header = xlNo
        .Orientation = xlSortColumns
        .MatchCase = False
        .SortMethod = xlPinYin
        .SetRange rng:=Selection
        .Apply
    End With
End Sub
Sub Macro2()
Attribute Macro2.VB_Description = "���� �ֵ��� ¼�ƣ�ʱ��: 2019/04/20"
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " 14"
'
' Macro2 Macro
' ���� �ֵ��� ¼�ƣ�ʱ��: 2019/04/20
'

'
    Selection.NumberFormatLocal = "@"
End Sub
