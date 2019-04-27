Attribute VB_Name = "ReportTests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod
Public Sub GetDispTbCrossReferenceItem_Tests()
    On Error GoTo TestFail
    
    'Arrange:
    Dim r1 As Integer 'result1
    Dim r2 As Integer
    Dim GlobalWC(1 To 10) As Integer
    Dim StrainGlobalWC(1 To 10) As Integer
    Dim strainNWCs As Integer
    Dim t As New ReportService
    
    Set t = New ReportService
    GlobalWC(1) = 1: GlobalWC(2) = 2: GlobalWC(3) = 3
    StrainGlobalWC(1) = 1: StrainGlobalWC(2) = 1: StrainGlobalWC(3) = 2

    'Act:
    r1 = t.GetDispTbCrossReferenceItem(1, GlobalWC, StrainGlobalWC, 3)
    r2 = t.GetDispTbCrossReferenceItem(2, GlobalWC, StrainGlobalWC, 3)
    'Assert:
    Assert.AreEqual 1, r1, "Not Equal"
    Assert.AreEqual 4, r2, "Not Equal"

     Set t = Nothing
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub GetStrainTbCrossReferenceItem_Tests()
    On Error GoTo TestFail
    
    'Arrange:
    Dim r1 As Integer 'result1
    Dim r2 As Integer
    Dim r3 As Integer
    Dim GlobalWC(1 To 10) As Integer
    Dim StrainGlobalWC(1 To 10) As Integer
    Dim NWCs As Integer
    Dim t As New ReportService
    
    Set t = New ReportService
    GlobalWC(1) = 1: GlobalWC(2) = 2: GlobalWC(3) = 3
    StrainGlobalWC(1) = 1: StrainGlobalWC(2) = 1: StrainGlobalWC(3) = 2

    'Act:
    r1 = t.GetStrainTbCrossReferenceItem(1, GlobalWC, 3, StrainGlobalWC)
    r2 = t.GetStrainTbCrossReferenceItem(2, GlobalWC, 3, StrainGlobalWC)
    r3 = t.GetStrainTbCrossReferenceItem(3, GlobalWC, 3, StrainGlobalWC)
    'Assert:
    Assert.AreEqual 2, r1, "Not Equal"
    Assert.AreEqual 3, r2, "Not Equal"
    Assert.AreEqual 5, r3, "Not Equal"
     Set t = Nothing
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

