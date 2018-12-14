Attribute VB_Name = "OfficialTestExample"
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
Public Sub AcceptsNumericKeys()
    On Error GoTo TestFail
    
    'Arrange:
    Dim value As String
    value = vbNullString
    
    Dim sut As NumKeyValidator 'sut denotes SystemUnderTest
    Set sut = New NumKeyValidator

    'Act:
    'Assert:
    Dim i As Integer
    Dim testResult As Boolean
    For i = 0 To 9
        testResult = sut.IsValidKey(Asc(CStr(i)), value)
        If Not testResult Then GoTo TestExit ' Exit if any test fails
    Next
    
TestExit:
    Assert.IsTrue testResult, "Value '" & i & "' was not accepted."
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


