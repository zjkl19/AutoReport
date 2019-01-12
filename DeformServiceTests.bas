Attribute VB_Name = "DeformServiceTests"
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
Public Sub GetStrain()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As DeformService 'sut denotes SystemUnderTest
    Set sut = New DeformService

    'Act:
    'Assert:
    Dim expectedResult As String
    expectedResult = "9.71"
 
    'Assert.AreEqual expected, actual, "oops, didn't expect this"
    Assert.AreEqual expectedResult, Format(sut.GetStrain(802.5, 800.5, 25.3, 24.1), "Fixed"), "Not Equal"
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub AutoDisp_Click_Tests()
    On Error GoTo TestFail
    
    'Arrange:
    Sheets("�ӶȲ���").Activate
    'Act:
    AutoDisp_Click
    'Assert:
    
    Assert.AreEqual "2.91", CStr(Format(Cells(20 + 5, 5), "Fixed")), "Not Equal" '�ܱ���
    Assert.AreEqual "0.14", CStr(Format(Cells(20 + 5, 8), "Fixed")), "Not Equal" '�������
    Assert.AreEqual "2.77", CStr(Format(Cells(20 + 5, 9), "Fixed")), "Not Equal" '���Ա���
    Assert.AreEqual "0.41", CStr(Format(Cells(20 + 5, 11), "Fixed")), "Not Equal" 'У��ϵ��
    Assert.AreEqual "4.81%", CStr(Format(Cells(20 + 5, 12), "Percent")), "Not Equal" '��Բ������
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub AutoStrain_Click_Tests()
    On Error GoTo TestFail
    
    'Arrange:
    Sheets("Ӧ�����").Activate
    'Act:
    AutoStrain_Click
    'Assert:
    
    Assert.AreEqual "43.22", CStr(Format(Cells(21 + 5, 27), "Fixed")), "Not Equal" '��Ӧ��
    Assert.AreEqual "42.07", CStr(Format(Cells(21 + 5, 28), "Fixed")), "Not Equal" '����Ӧ��
    Assert.AreEqual "1.15", CStr(Format(Cells(21 + 5, 29), "Fixed")), "Not Equal" '����Ӧ��
    Assert.AreEqual "0.40", CStr(Format(Cells(21 + 5, 31), "Fixed")), "Not Equal" 'У��ϵ��
    Assert.AreEqual "2.66%", CStr(Format(Cells(21 + 5, 32), "Percent")), "Not Equal" '��Բ���Ӧ��
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
