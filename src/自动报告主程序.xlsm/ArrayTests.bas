Attribute VB_Name = "ArrayTests"
Option Explicit

Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod
Public Sub CountElements_Tests()
    On Error GoTo TestFail
    
    'Arrange:
    Dim cts1 As Integer 'counts:计数
    Dim cts2 As Integer
    Dim a(1 To 10) As Integer
    Dim t As New ArrayService
    Set t = New ArrayService
    a(1) = 1
    a(2) = 1
    a(3) = 1
    a(4) = 2
    'Act:
    cts1 = t.CountElements(a, 1)
    cts2 = t.CountElements(a, 1, 2)
    'Assert:
    Assert.AreEqual 3, cts1, "Not Equal"
    Assert.AreEqual 2, cts2, "Not Equal"

     Set t = Nothing
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub ArrayMax_Tests()
    On Error GoTo TestFail
    
    'Arrange:
    Dim max1 As Integer 'max:所求最大值结果
    Dim max2 As Integer
    Dim a(1 To 10) As Integer
    Dim t As New ArrayService
    Set t = New ArrayService
    a(1) = 1
    a(2) = 1
    a(3) = 2
    a(4) = 2
    a(5) = 3
    a(6) = 3
    'Act:
    max1 = t.ArrayMax(a)
    max2 = t.ArrayMax(a, 4)
    'Assert:
    Assert.AreEqual 3, max1, "Not Equal"
    Assert.AreEqual 2, max2, "Not Equal"

     Set t = Nothing
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

