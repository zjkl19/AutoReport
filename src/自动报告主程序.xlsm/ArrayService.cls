VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'统计一维整形数组某个元素个数
'myArray：一维数组
'elementMatch：欲匹配的元素
'uB:查找的上界（下界为1）
'举例：A=(1,1,2,2,3)
'CountElements(A,1)==2
'CountElements(A,2,3)==1

Public Function CountElements(ByRef myArray() As Integer, ByVal elementMatch As Integer, Optional uB) As Integer
        Dim i As Integer
        Dim uBoundSet As Integer    '设置的查找上界
        Dim counts As Integer    '总数
        counts = 0    '默认为0个
        
        '没有手动输入查找上界则采用数组的上界
        '注意：若使用Optional uB as integer，则uB不传参时值为0
        If IsMissing(uB) Then
              uBoundSet = UBound(myArray)
        Else
              uBoundSet = uB
        End If
            
        For i = 1 To uBoundSet
            If myArray(i) = elementMatch Then
                counts = counts + 1
            End If
        Next i
        CountElements = counts

End Function

'求一维整形数组最大值
'myArray：一维数组
'uB:查找的上界（下界为1）
'举例：A=(1,1,2,2,3)
'ArrayMax(A)==3
'ArrayMax(A,3)==2

Public Function ArrayMax(ByRef myArray() As Integer, Optional uB) As Integer
        Dim i As Integer
        Dim uBoundSet As Integer    '设置的查找上界
        Dim result As Integer
        result = myArray(1)
        
        '没有手动输入查找上界则采用数组的上界
        '注意：若使用Optional uB as integer，则uB不传参时值为0
        If IsMissing(uB) Then
              uBoundSet = UBound(myArray)
        Else
              uBoundSet = uB
        End If
            
        For i = 2 To uBoundSet
            If result < myArray(i) Then
                result = myArray(i)
            End If
        Next i
        ArrayMax = result

End Function
