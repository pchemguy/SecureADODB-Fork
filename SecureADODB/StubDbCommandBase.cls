VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StubDbCommandBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A stub acting as a IDbCommandBase implementation."
'@Folder "SecureADODB.DbCommand. Tests"
'@ModuleDescription("A stub acting as a IDbCommandBase implementation.")
'@IgnoreModule ProcedureNotUsed
Option Explicit

Implements IDbCommandBase

Private Type TInvokeState
    CreateCommandInvokes As Long
    ValidateOrdinalArgumentsInvokes As Long
End Type
Private this As TInvokeState


Public Property Get CreateCommandInvokes() As Long
    CreateCommandInvokes = this.CreateCommandInvokes
End Property

Public Property Get ValidateOrdinalArgumentsInvokes() As Long
    ValidateOrdinalArgumentsInvokes = this.ValidateOrdinalArgumentsInvokes
End Property

Private Property Get IDbCommandBase_Self() As IDbCommandBase
    Set IDbCommandBase_Self = Me
End Property

'@Description("Returns class reference.")
Private Property Get IDbCommandBase_Class() As Object
    Set IDbCommandBase_Class = DbCommandBase
End Property

Private Function IDbCommandBase_CreateCommand(ByVal db As IDbConnection, ByVal CommandType As ADODB.CommandTypeEnum, ByVal sql As String, args() As Variant) As ADODB.Command
    this.CreateCommandInvokes = this.CreateCommandInvokes + 1
    Dim result As ADODB.Command
    Set result = New ADODB.Command
    result.CommandType = CommandType
    result.CommandText = sql
    Set IDbCommandBase_CreateCommand = result
End Function

