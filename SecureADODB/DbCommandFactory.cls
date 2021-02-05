VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbCommandFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An implementation of an abstract factory that creates DbCommand objects."
'@Folder "SecureADODBmod.DbCommand"
'@IgnoreModule UseMeaningfulName, IllegalAnnotation
'@PredeclaredId
'@ModuleDescription("An implementation of an abstract factory that creates DbCommand objects.")

Option Explicit

Implements IDbCommandFactory

Private Type TFactory
    commandClass As Object
    Base As IDbCommandBase
End Type
Private this As TFactory


'@Description("Default abstract factory")
'@Ignore ProcedureNotUsed
Public Function Create(ByVal baseCommand As IDbCommandBase, Optional ByVal commandClassName As String = "DefaultDbCommand") As IDbCommandFactory
Attribute Create.VB_Description = "Default abstract factory"
    Dim result As DbCommandFactory
    Set result = New DbCommandFactory
    result.CreInit baseCommand, commandClassName
    Set Create = result
End Function


'@Description("Default constructor")
Public Sub CreInit(ByVal baseCommand As IDbCommandBase, Optional ByVal commandClassName As String = "DefaultDbCommand")
Attribute CreInit.VB_Description = "Default constructor"
    Set this.Base = baseCommand
    Select Case commandClassName
        Case "DefaultDbCommand"
            Set this.commandClass = DefaultDbCommand
        Case "DbCommand"
            Set this.commandClass = DbCommand
        Case Else
            Dim errorDetails As TError
            With errorDetails
                .number = ErrNo.NoObject
                .Name = "NoObject"
                .source = "DbCommandFactory"
                .message = "Unknown DbCommand class: " & commandClassName
                .description = .message
            End With
            RaiseError errorDetails
    End Select
End Sub


'@Description("Default interface factory")
Private Function IDbCommandFactory_CreateInstance(ByVal db As IDbConnection, Optional ByVal UseDisconnectedRecordset As Boolean = False) As IDbCommand
Attribute IDbCommandFactory_CreateInstance.VB_Description = "Default interface factory"
    Set IDbCommandFactory_CreateInstance = this.commandClass.Create(db, this.Base, UseDisconnectedRecordset)
End Function
