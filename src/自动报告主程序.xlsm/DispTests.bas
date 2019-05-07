Attribute VB_Name = "DispTests"
Option Explicit

Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

Private t As Disp    'Test Object

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
    
    'Arrange
    Set t = New Disp
    Dim NWCs As Integer
    Dim DispUbound(1 To 10) As Integer
    Dim GlobalWC(1 To MAX_NWC)  As Integer   '全局工况定位数组
    Dim GroupName(1 To MAX_NWC)  As String '每个组的名称
    Dim NodeName(1 To MAX_NWC, 1 To MAX_NPS) As String
    
    Dim i As Integer
    NWCs = 2
    DispUbound(1) = 4: DispUbound(2) = 6
    GlobalWC(1) = 1: GlobalWC(2) = 2
    GroupName(1) = "A-A": GroupName(2) = "A-A"
    NodeName(1, 1) = "A-1": NodeName(1, 2) = "A-2": NodeName(1, 3) = "A-3": NodeName(1, 4) = "A-4"
    NodeName(2, 1) = "A-4": NodeName(2, 2) = "A-5": NodeName(2, 3) = "A-6": NodeName(2, 4) = "A-7": NodeName(2, 5) = "A-8": NodeName(2, 6) = "A-9"
    'Act
    t.Init NWCs:=NWCs, DispUbound:=DispUbound, GlobalWC:=GlobalWC, GroupName:=GroupName, NodeName:=NodeName
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
    Set t = Nothing
End Sub

'@TestMethod
'基本参数可以传入类中
Public Sub Can_Invoke_Argu()
    On Error GoTo TestFail
    
    'Arrange:
     
    'Act:
    
    'Assert:
    Assert.AreEqual 2, t.NWCs, "Not Equal"
    
    Assert.AreEqual 1, t.GlobalWC(1), "Not Equal"
    Assert.AreEqual 2, t.GlobalWC(2), "Not Equal"
    
    Assert.AreEqual "A-A", t.GroupName(1), "Not Equal"
    Assert.AreEqual "A-A", t.GroupName(2), "Not Equal"
    
    Assert.AreEqual 4, t.DispUbound(1), "Not Equal"
    Assert.AreEqual 6, t.DispUbound(2), "Not Equal"
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
'可以传入试验值
Public Sub Can_Invoke_ExpValue()
    On Error GoTo TestFail
    
    'Arrange:
     
    'Act:
    
    'Assert:
    Assert.AreEqual "A-2", t.NodeName(1, 2), "Not Equal"
    Assert.AreEqual "A-6", t.NodeName(2, 3), "Not Equal"

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



