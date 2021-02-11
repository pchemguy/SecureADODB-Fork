VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An object that builds and executes a parameterized ADODB command from a command string and parameter values. Returns connected recordsets that the client should close or disconnected recordset."
'@Folder "SecureADODB.DbCommand"
'@ModuleDescription("An object that builds and executes a parameterized ADODB command from a command string and parameter values. Returns connected recordsets that the client should close or disconnected recordset.")
'@PredeclaredId
'@IgnoreModule IllegalAnnotation
Option Explicit

Implements IDbCommand

Private Type TCommand
    cmd As ADODB.Command
    Base As IDbCommandBase
    Connection As IDbConnection
End Type
Private this As TCommand


'@Inject: just an idea.. see #https://github.com/rubberduck-vba/Rubberduck/issues/5463
Public Function Create(ByVal db As IDbConnection, ByVal baseCommand As IDbCommandBase) As IDbCommand
    Errors.GuardNonDefaultInstance Me, DbCommand
    
    Dim result As DbCommand
    Set result = New DbCommand
    result.CreInit db, baseCommand
    Set Create = result
End Function


'@Description("Default constructor")
Friend Sub CreInit(ByVal db As IDbConnection, ByVal baseCommand As IDbCommandBase)
Attribute CreInit.VB_Description = "Default constructor"
    Errors.GuardNullReference db
    Errors.GuardNullReference baseCommand

    Set this.Connection = db
    Set this.Base = baseCommand
End Sub

Public Function IDbCommand_AdoCommand(ByVal sql As String, ParamArray args()) As ADODB.Command
    Dim localArgs() As Variant
    '''' Recordset routines pass their (sql, ParamArray) for processing here.
    '''' If the first item in args is also an array, assume ParamArray to ParamArray
    '''' delegation. Should unfold the outer array.
    If IsArray(args(0)) Then
        localArgs = args(0)
    Else
        localArgs = args
    End If
    
    Set this.cmd = this.Base.CreateCommand(this.Connection, adCmdText, sql, localArgs)
    Set IDbCommand_AdoCommand = this.cmd
End Function


Private Function IDbCommand_Execute(ByVal sql As String, ParamArray args() As Variant) As ADODB.Recordset
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant
    localArgs = args
    
    Set this.cmd = this.Base.CreateCommand(this.Connection, adCmdText, sql, localArgs)
    If True Then
        Set IDbCommand_Execute = this.Base.GetDisconnectedRecordset(this.cmd)
    Else
        Set IDbCommand_Execute = this.cmd.Execute
    End If
End Function


Private Sub IDbCommand_ExecuteNonQuery(ByVal sql As String, ParamArray args() As Variant)
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant
    localArgs = args
    
    Set this.cmd = this.Base.CreateCommand(this.Connection, adCmdText, sql, localArgs)
    '@Ignore FunctionReturnValueDiscarded
    this.cmd.Execute
End Sub


Private Function IDbCommand_ExecuteScalar(ByVal sql As String, ParamArray args() As Variant) As Variant
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant
    localArgs = args
    
    IDbCommand_ExecuteScalar = this.Base.ExecuteScalar(this.Connection, adCmdText, sql, localArgs)
End Function
