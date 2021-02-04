VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefaultDbCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An object that builds and executes a parameterized ADODB command from a command string and parameter values. Returns connected recordsets that the client should close."
'@Folder "SecureADODB.DbCommand"
'@ModuleDescription("An object that builds and executes a parameterized ADODB command from a command string and parameter values. Returns connected recordsets that the client should close.")
'@PredeclaredId
'@IgnoreModule UseMeaningfulName, IllegalAnnotation
Option Explicit

Implements IDbCommand

Private Type TCommand
    Base As IDbCommandBase
    Connection As IDbConnection
End Type
Private this As TCommand


'@Inject: just an idea.. see #https://github.com/rubberduck-vba/Rubberduck/issues/5463
Public Function Create(ByVal db As IDbConnection, ByVal baseCommand As IDbCommandBase) As IDbCommand
    Errors.GuardNonDefaultInstance Me, DefaultDbCommand
    
    Dim result As DefaultDbCommand
    Set result = New DefaultDbCommand
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


Private Sub GuardInvalidOrdinalArguments(ByVal sql As String, ByRef args() As Variant)
    Errors.GuardExpression Not this.Base.ValidateOrdinalArguments(sql, args), message:="Arguments provided are inconsistent with expected ordinal parameters."
End Sub


Private Function IDbCommand_Execute(ByVal sql As String, ParamArray args() As Variant) As ADODB.Recordset
    Errors.GuardNullReference this.Connection
    Errors.GuardExpression this.Connection.State <> adStateOpen, message:="Connection is not open."
    
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant: localArgs = args
    GuardInvalidOrdinalArguments sql, localArgs
    
    Dim cmd As ADODB.Command
    Set cmd = this.Base.CreateCommand(this.Connection, adCmdText, sql, localArgs)
    
    Set IDbCommand_Execute = cmd.Execute
End Function


Private Sub IDbCommand_ExecuteNonQuery(ByVal sql As String, ParamArray args() As Variant)
    Errors.GuardNullReference this.Connection
    Errors.GuardExpression this.Connection.State <> adStateOpen, message:="Connection is not open."
    
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant: localArgs = args
    GuardInvalidOrdinalArguments sql, localArgs
    
    Dim cmd As ADODB.Command
    Set cmd = this.Base.CreateCommand(this.Connection, adCmdText, sql, localArgs)
    '@Ignore FunctionReturnValueDiscarded
    cmd.Execute
End Sub


Private Function IDbCommand_ExecuteWithParameters(ByVal sql As String, ParamArray params() As Variant) As ADODB.Recordset
    Errors.GuardNullReference this.Connection
    Errors.GuardExpression this.Connection.State <> adStateOpen, message:="Connection is not open."
    
    Dim localParams() As Variant
    localParams = params
    GuardInvalidOrdinalArguments sql, localParams
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    Set cmd.ActiveConnection = this.Connection.AdoConnection
    cmd.commandType = adCmdText
    cmd.CommandText = sql
    
    Dim i As Long
    For i = LBound(params) To UBound(params)
        cmd.Parameters.Append params(i)
    Next
    
    Dim result As ADODB.Recordset
    Set result = cmd.Execute
    
    Set IDbCommand_ExecuteWithParameters = result
End Function


Private Function IDbCommand_GetSingleValue(ByVal sql As String, ParamArray args() As Variant) As Variant
    Errors.GuardNullReference this.Connection
    Errors.GuardExpression this.Connection.State <> adStateOpen, message:="Connection is not open."
    
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant: localArgs = args
    GuardInvalidOrdinalArguments sql, localArgs
    
    IDbCommand_GetSingleValue = this.Base.GetSingleValue(this.Connection, sql, localArgs)
End Function
