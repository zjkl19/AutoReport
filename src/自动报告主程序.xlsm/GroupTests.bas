Attribute VB_Name = "GroupTests"
Option Explicit
'分组有关功能测试

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

'测试组数的获取是否正确
'@TestMethod
Public Sub GetGroupCounts_Tests()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gt As GroupService
    Dim GroupName(1 To 100, 1 To 10) As String
    
    Set gt = New GroupService
    GroupName(1, 1) = "A-A"
    GroupName(1, 2) = "A-A"
    GroupName(1, 3) = "B-B"
    GroupName(1, 4) = "B-B"
    
    GroupName(2, 1) = "A-A"
    GroupName(2, 2) = "A-A"
    GroupName(2, 3) = "B-B"
    GroupName(2, 4) = "B-B"
    GroupName(2, 5) = "C-C"
    GroupName(2, 6) = "C-C"
    
    'Act:
    'Assert:
    Assert.AreEqual 2, gt.GetGroupCounts(GroupName, 1, 4), "Not Equal"
    Assert.AreEqual 3, gt.GetGroupCounts(GroupName, 2, 6), "Not Equal"
     Set gt = Nothing
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'测试组名的获取是否正确
'@TestMethod
Public Sub GetGroupName_Tests()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gt As GroupService
    Dim GroupName(1 To 100, 1 To 10) As String
    
    Set gt = New GroupService
    GroupName(1, 1) = "A-A"
    GroupName(1, 2) = "A-A"
    GroupName(1, 3) = "B-B"
    GroupName(1, 4) = "B-B"
    
    GroupName(2, 1) = "A-A"
    GroupName(2, 2) = "A-A"
    GroupName(2, 3) = "B-B"
    GroupName(2, 4) = "B-B"
    GroupName(2, 5) = "C-C"
    GroupName(2, 6) = "C-C"
    
    'Act:
    'Assert:
    Assert.AreEqual "B-B", gt.GetGroupName(GroupName, 1, 2), "Not Equal"
    Assert.AreEqual "C-C", gt.GetGroupName(GroupName, 2, 3), "Not Equal"
     Set gt = Nothing
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'测试组索引的获取是否正确
'@TestMethod
Public Sub GetFirstAndLastIndex_Tests()
    On Error GoTo TestFail
    
    'Arrange:
    Dim gs As GroupService
    Dim GroupName(1 To 100, 1 To 10) As String
    Dim firstI As Integer
    Dim lastI As Integer
    Dim firstI2 As Integer
    Dim lastI2 As Integer
    
    Set gs = New GroupService
    GroupName(1, 1) = "A-A"
    GroupName(1, 2) = "A-A"
    GroupName(1, 3) = "B-B"
    GroupName(1, 4) = "B-B"
    
    GroupName(2, 1) = "A-A"
    GroupName(2, 2) = "A-A"
    GroupName(2, 3) = "B-B"
    GroupName(2, 4) = "B-B"
    GroupName(2, 5) = "C-C"
    GroupName(2, 6) = "C-C"
    
    'Act:
    gs.GetFirstAndLastIndex GroupName, 1, 2, 4, firstI, lastI
    gs.GetFirstAndLastIndex GroupName, 2, 2, 6, firstI2, lastI2
    'Assert:
    Assert.AreEqual 3, firstI, "Not Equal"
    Assert.AreEqual 4, lastI, "Not Equal"
    Assert.AreEqual 3, firstI, "Not Equal"
    Assert.AreEqual 4, lastI, "Not Equal"
     Set gs = Nothing
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
