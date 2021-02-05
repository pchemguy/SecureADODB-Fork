Attribute VB_Name = "SharedStructuresForTests"
'@Folder "SecureADODBmod.Guard.Tests"
'@TestModule
'@IgnoreModule VariableNotAssigned, UnassignedVariableUsage
Option Explicit
Option Private Module


Const MsgExpectedErrNotRaised As String = "Expected error was not raised."
Const MsgUnexpectedErrRaised As String = "Unexpected error was raised."


Public Sub AssertExpectedError(Optional ByVal ExpectedErrorNo As ErrNo = ErrNo.PassedNoErr)
    Dim ActualErrNo As Long
    ActualErrNo = VBA.Err.number
    Dim errorDetails As String
    errorDetails = " Error: #" & ActualErrNo & " - " & VBA.Err.description
    VBA.Err.Clear
    
    Dim Assert As Object
    Select Case ActualErrNo
        Case ExpectedErrorNo
            Assert.Succeed
        Case ErrNo.PassedNoErr
            Assert.Fail MsgExpectedErrNotRaised
        Case Else
            Assert.Fail MsgUnexpectedErrRaised & errorDetails
    End Select
End Sub

