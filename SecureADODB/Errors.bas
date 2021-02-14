Attribute VB_Name = "Errors"
Attribute VB_Description = "Global procedures for throwing common errors."
'@Folder "SecureADODB"
'@ModuleDescription("Global procedures for throwing common errors.")
Option Explicit
Option Private Module

Public Const SecureADODBCustomError As Long = vbObjectError Or 32


'@Description("Re-raises the current error, if there is one.")
Public Sub RethrowOnError()
Attribute RethrowOnError.VB_Description = "Re-raises the current error, if there is one."
    With VBA.Information.Err
        If .number <> 0 Then
            Debug.Print "Error " & .number, .description
            .Raise .number
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified Boolean expression is True.")
Public Sub GuardExpression(ByVal throw As Boolean, _
                  Optional ByVal source As String = "SecureADODB.Errors", _
                  Optional ByVal message As String = "Invalid procedure call or argument.")
Attribute GuardExpression.VB_Description = "Raises a run-time error if the specified Boolean expression is True."
    If throw Then VBA.Information.Err.Raise SecureADODBCustomError, source, message
End Sub

'@Description("Raises a run-time error if the specified instance isn't the default instance.")
Public Sub GuardNonDefaultInstance(ByVal instance As Object, ByVal DefaultInstance As Object, _
                          Optional ByVal source As String = "SecureADODB.Errors", _
                          Optional ByVal message As String = "Method should be invoked from the default/predeclared instance of this class.")
Attribute GuardNonDefaultInstance.VB_Description = "Raises a run-time error if the specified instance isn't the default instance."
    Debug.Assert TypeName(instance) = TypeName(DefaultInstance)
    GuardExpression Not instance Is DefaultInstance, source, message
End Sub

'@Description("Raises a run-time error if the specified object reference is already set.")
Public Sub GuardDoubleInitialization(ByVal instance As Object, _
                            Optional ByVal source As String = "SecureADODB.Errors", _
                            Optional ByVal message As String = "Object is already initialized.")
Attribute GuardDoubleInitialization.VB_Description = "Raises a run-time error if the specified object reference is already set."
    GuardExpression Not instance Is Nothing, source, message
End Sub

'@Description("Raises a run-time error if the specified object reference is Nothing.")
Public Sub GuardNullReference(ByVal instance As Object, _
                     Optional ByVal source As String = "SecureADODB.Errors", _
                     Optional ByVal message As String = "Object reference cannot be Nothing.")
Attribute GuardNullReference.VB_Description = "Raises a run-time error if the specified object reference is Nothing."
    GuardExpression instance Is Nothing, source, message
End Sub

'@Description("Raises a run-time error if the specified string is empty.")
Public Sub GuardEmptyString(ByVal value As String, _
                   Optional ByVal source As String = "SecureADODB.Errors", _
                   Optional ByVal message As String = "String cannot be empty.")
Attribute GuardEmptyString.VB_Description = "Raises a run-time error if the specified string is empty."
    GuardExpression value = vbNullString, source, message
End Sub
