Attribute VB_Name = "DbConnectionTests"
'@Folder "SecureADODB.DbConnection.Tests"
'@TestModule
'@IgnoreModule
Option Explicit
Option Private Module

Private Const ERR_INVALID_WITHOUT_LIVE_CONNECTION = 3709

#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If


'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub


Private Function GetTestConnectionString() As String
    GetTestConnectionString = "connection string"
End Function


'@TestMethod("Factory Guard")
Private Sub Create_ThrowsIfNotInvokedFromDefaultInstance()
    On Error GoTo TestFail
    
    With New DbConnection
        On Error GoTo CleanFail
        Dim sut As DbConnection
        Set sut = .Create(GetTestConnectionString)
        On Error GoTo 0
    End With
CleanFail:
    If Err.Number = ErrNo.NonDefaultInstanceErr Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub


'@TestMethod("Factory Guard")
Private Sub Create_ThrowsWithEmptyConnectionString()
    On Error GoTo TestFail
    
    With New DbConnection
        On Error GoTo CleanFail
        Dim sut As DbConnection
        Set sut = .Create(vbNullString)
        On Error GoTo 0
    End With
CleanFail:
    If Err.Number = ErrNo.NonDefaultInstanceErr Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub
