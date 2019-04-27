VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GroupService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'获取某特定工况组数
'groupName(i,j)：第i个工况，第j个测点组名
'wc：工况号
'nP:测点数
Public Function GetGroupCounts(ByRef GroupName() As String, ByVal wc As Integer, ByVal nP As Integer) As Integer
        Dim i As Integer
        Dim counts As Integer    '总数
        counts = 0
        Dim currName As String
        currName = ""
        
        For i = 1 To nP
            If GroupName(wc, i) <> currName Then
                counts = counts + 1
                currName = GroupName(wc, i)
            End If
        Next i
        GetGroupCounts = counts

End Function

'获取第i个工况组数
'组数不能越界
'groupName(i,j)：第i个工况，第j个测点组名
'wc：工况号
'gn:组号
Public Function GetGroupName(ByRef GroupName() As String, ByVal wc As Integer, ByVal gn As Integer) As String
        Dim i As Integer
        i = 1
        Dim currGroup, currPointer As Integer
        currGroup = 1
        currPointer = 1
        
        Dim gName As String
        gName = GroupName(wc, 1)    '默认为该工况第1个测点的组名
        
        Dim currName As String
        currName = GroupName(wc, 1)

        '举例识别算法
        While currGroup < gn
            If GroupName(wc, currPointer) = currName Then
                 currPointer = currPointer + 1
           Else
                  currName = GroupName(wc, currPointer)
                  currGroup = currGroup + 1
                  If currGroup < gn Then     '尚未找到
                        currPointer = currPointer + 1
                  End If
            End If
        Wend
        GetGroupName = GroupName(wc, currPointer)

End Function

'获取第i个工况第i组首尾元素索引
'组数不能越界
'groupName(i,j)：第i个工况，第j个测点组名
'wc：工况号
'gn:组号
'np:该工况测点数
'firstI:首个点索引
'lastI:最后一个点索引
Public Sub GetFirstAndLastIndex(ByRef GroupName() As String, ByVal wc As Integer, ByVal gn As Integer, ByVal nP As Integer, ByRef firstI As Integer, ByRef lastI As Integer)

        firstI = 1: lastI = 1    '首尾元素初始化为1
        
        Dim i As Integer
        i = 1
        Dim currGroup, currPointer As Integer
        currGroup = 1
        currPointer = 1
        
        Dim gName As String
        gName = GroupName(wc, 1)    '默认为该工况第1个测点的组名
        
        Dim currName As String
        currName = GroupName(wc, 1)

        '举例识别算法
        While currGroup < gn
            If GroupName(wc, currPointer) = currName Then
                 currPointer = currPointer + 1
           Else
                  currName = GroupName(wc, currPointer)
                  currGroup = currGroup + 1
                  If currGroup < gn Then     '尚未找到
                        currPointer = currPointer + 1
                  End If
            End If
        Wend
         firstI = currPointer
         
        While GroupName(wc, currPointer) = currName And currPointer < nP
                 currPointer = currPointer + 1
        Wend
        If GroupName(wc, currPointer) <> currName Then
            lastI = currPointer - 1
        Else
            lastI = currPointer
        End If

End Sub
