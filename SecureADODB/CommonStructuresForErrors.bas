Attribute VB_Name = "CommonStructuresForErrors"
'@Folder("Guard")
Option Explicit


Public Enum ErrNo
    PassedNoErr = 0&
    TypeMismatchErr = 13&
    FileNotFoundErr = 53&
    ObjectNotSetErr = 91&
    ObjectRequiredErr = 424&
    InvalidObjectUseErr = 425&
    MemberNotExistErr = 438&
    ActionNotSupportedErr = 445&
    
    CustomErr = VBA.vbObjectError + 1000&
    NotImplementedErr = VBA.vbObjectError + 1001&
    DefaultInstanceErr = VBA.vbObjectError + 1011&
    NonDefaultInstanceErr = VBA.vbObjectError + 1012&
    EmptyStringErr = VBA.vbObjectError + 1013&
    SingletonErr = VBA.vbObjectError + 1014&
    UnknownClassErr = VBA.vbObjectError + 1015&
    ObjectSetErr = VBA.vbObjectError + 1091&
End Enum


Public Type TError
    number As ErrNo
    name As String
    source As String
    message As String
    description As String
    trapped As Boolean
End Type


'@Ignore ProcedureNotUsed
'@Description("Re-raises the current error, if there is one.")
Public Sub RethrowOnError()
Attribute RethrowOnError.VB_Description = "Re-raises the current error, if there is one."
    With VBA.Err
        If .number <> 0 Then
            Debug.Print "Error " & .number, .description
            .Raise .number
        End If
    End With
End Sub


'@Description("Formats and raises a run-time error.")
Public Sub RaiseError(ByRef errorDetails As TError)
Attribute RaiseError.VB_Description = "Formats and raises a run-time error."
    With errorDetails
        Dim message As Variant
        message = Array("Error:", _
            "name: " & .name, _
            "number: " & .number, _
            "message: " & .message, _
            "description: " & .description, _
            "source: " & .source)
        Debug.Print Join(message, vbNewLine & vbTab)
        VBA.Err.Raise .number, .source, .message
    End With
End Sub
