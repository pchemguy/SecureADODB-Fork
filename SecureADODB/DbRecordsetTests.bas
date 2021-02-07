Attribute VB_Name = "DbRecordsetTests"
Attribute VB_Description = "Tests for the DbRecordset class."
'@Folder "SecureADODB.DbRecordset"
''@TestModule
'@ModuleDescription("Tests for the DbRecordset class.")
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module


#Const LateBind = LateBindTests
#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If


'This method runs once per module.
'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Guard = Nothing
End Sub


'This method runs after every test in the module.
'@TestCleanup
Private Sub TestCleanup()
    Err.Clear
End Sub

'==================================================
'==================================================


'@TestMethod("Guard.Class")
Private Sub Class_CheckAvailability()
    On Error GoTo TestFail
    
Arrange:
    Dim classVar As Object: Set classVar = DbRecordset
Act:
    Dim classVarReturned As Object: Set classVarReturned = classVar.Create.Class
Assert:
    Assert.AreEqual TypeName(classVar), TypeName(classVarReturned), "Error: type mismatch: " & TypeName(classVarReturned) & " type."
    Assert.AreSame classVar, classVarReturned, "Error: bad Class pointer"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub
