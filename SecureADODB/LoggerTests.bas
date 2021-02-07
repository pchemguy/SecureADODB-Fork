Attribute VB_Name = "LoggerTests"
Attribute VB_Description = "Tests for the Logger class."
'@Folder "SecureADODB.Logger"
'@TestModule
'@ModuleDescription("Tests for the Logger class.")
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
End Sub


'This method runs after every test in the module.
'@TestCleanup
Private Sub TestCleanup()
    Err.Clear
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='

'@TestMethod("Factory Guard")
Private Sub ztcCreate_PassesIfInvokedFromDefaultInstance()
    On Error Resume Next
    Dim AdoLogger As ILogger: Set AdoLogger = Logger.Create
    AssertExpectedError Assert, ErrNo.PassedNoErr
End Sub


'@TestMethod("Factory Guard")
Private Sub ztcCreate_ThrowsIfNotInvokedFromDefaultInstance()
    On Error Resume Next
    Dim stubLogger As Logger: Set stubLogger = New Logger
    Dim stubILogger As Logger: Set stubILogger = stubLogger.Create
    Assert.IsNothing stubILogger
    AssertExpectedError Assert, ErrNo.NonDefaultInstanceErr
End Sub


'@TestMethod("Self")
Private Sub ztcSelf_CheckAvailability()
    On Error GoTo TestFail
    
Arrange:
    Dim instanceVar As Object: Set instanceVar = Logger.Create
Act:
    Dim selfVar As Object: Set selfVar = instanceVar.Self
Assert:
    Assert.AreEqual TypeName(instanceVar), TypeName(selfVar), "Error: type mismatch: " & TypeName(selfVar) & " type."
    Assert.AreSame instanceVar, selfVar, "Error: bad Self pointer"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("Class")
Private Sub ztcClass_CheckAvailability()
    On Error GoTo TestFail
    
Arrange:
    Dim classVar As Object: Set classVar = Logger
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


'===================================================='
'============== INTERACTIVE TEST CASES =============='
'===================================================='

'@TestMethod("PrintLog")
'@Description "Log some items, check the contents"
Private Sub ziPrintLogTest()
Attribute ziPrintLogTest.VB_Description = "Log some items, check the contents"
    Dim AdoLogger As ILogger
    Set AdoLogger = Logger.Create
    
    AdoLogger.Log "AAA"
    AdoLogger.Log "BBB"
    AdoLogger.PrintLog
End Sub
