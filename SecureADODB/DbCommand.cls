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
'@Folder "SecureADODBmod.DbCommand"
'@ModuleDescription("An object that builds and executes a parameterized ADODB command from a command string and parameter values. Returns connected recordsets that the client should close or disconnected recordset.")
'@PredeclaredId
'@IgnoreModule UseMeaningfulName, IllegalAnnotation
Option Explicit

Implements IDbCommand

Private Type TCommand
    Base As IDbCommandBase
    Connection As IDbConnection
    UseDisconnectedRecordset As Boolean
End Type
Private this As TCommand


'@Inject: just an idea.. see #https://github.com/rubberduck-vba/Rubberduck/issues/5463
Public Function Create(ByVal db As IDbConnection, _
                       ByVal baseCommand As IDbCommandBase, _
              Optional ByVal UseDisconnectedRecordset As Boolean = False) As IDbCommand
              
    Errors.GuardNonDefaultInstance Me, DbCommand
    
    Dim result As DbCommand
    Set result = New DbCommand
    result.CreInit db, baseCommand, UseDisconnectedRecordset
    Set Create = result
End Function


'@Description("Default constructor")
Friend Sub CreInit(ByVal db As IDbConnection, _
                       ByVal baseCommand As IDbCommandBase, _
              Optional ByVal UseDisconnectedRecordset As Boolean = False)
Attribute CreInit.VB_Description = "Default constructor"
              
    Errors.GuardNullReference db
    Errors.GuardNullReference baseCommand

    Set this.Connection = db
    Set this.Base = baseCommand
    this.UseDisconnectedRecordset = UseDisconnectedRecordset
End Sub


Private Function IDbCommand_Execute(ByVal sql As String, ParamArray args() As Variant) As ADODB.Recordset
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant
    localArgs = args
    
    Dim cmd As ADODB.Command
    Set cmd = this.Base.CreateCommand(this.Connection, adCmdText, sql, localArgs)
    Set IDbCommand_Execute = cmd.Execute
End Function


Private Function IDbCommand_ExecuteWithParameters(ByVal sql As String, ParamArray params() As Variant) As ADODB.Recordset
    Dim localParams() As Variant
    localParams = params
        
    Dim emptyArgs() As Variant
    Dim cmd As ADODB.Command
    Set cmd = this.Base.CreateCommand(this.Connection, adCmdText, sql, emptyArgs)
    
    Dim i As Long
    For i = LBound(localParams) To UBound(localParams)
        cmd.Parameters.Append localParams(i)
    Next
    
    Set IDbCommand_ExecuteWithParameters = cmd.Execute
End Function


Private Sub IDbCommand_ExecuteNonQuery(ByVal sql As String, ParamArray args() As Variant)
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant
    localArgs = args
    
    Dim cmd As ADODB.Command
    Set cmd = this.Base.CreateCommand(this.Connection, adCmdText, sql, localArgs)
    '@Ignore FunctionReturnValueDiscarded
    cmd.Execute
End Sub


Private Function IDbCommand_GetSingleValue(ByVal sql As String, ParamArray args() As Variant) As Variant
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant
    localArgs = args
    
    IDbCommand_GetSingleValue = this.Base.GetSingleValue(this.Connection, adCmdText, sql, localArgs)
End Function
