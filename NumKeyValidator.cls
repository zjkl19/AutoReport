VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NumKeyValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function IsValidKey(ByVal keyAscii As Integer, ByVal value As String) As Boolean
    IsValidKey = keyAscii >= KeyCodeConstants.vbKey0 And keyAscii <= KeyCodeConstants.vbKey9
End Function

