VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoDbCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An object that deals with the database connection and builds and executes a parameterized ADODB command from a command string and parameter values. Returns disconnected recordsets."
'@Folder "SecureADODB.DbCommand"
'@ModuleDescription("An object that deals with the database connection and builds and executes a parameterized ADODB command from a command string and parameter values. Returns disconnected recordsets.")
'@PredeclaredId
'@IgnoreModule UseMeaningfulName, IllegalAnnotation
Option Explicit

Implements IDbCommand

Private Type TCommand
    ConnectionString As String
    ConnectionFactory As IDbConnectionFactory
    Base As IDbCommandBase
End Type
Private this As TCommand


'@Inject: just an idea.. see #https://github.com/rubberduck-vba/Rubberduck/issues/5463
Public Function Create(ByVal connString As String, ByVal factory As IDbConnectionFactory, ByVal baseCommand As IDbCommandBase) As IDbCommand
    Errors.GuardNonDefaultInstance Me, AutoDbCommand
    
    Dim result As AutoDbCommand
    Set result = New AutoDbCommand
    result.CreInit connString, factory, baseCommand
    Set Create = result
End Function


'@Description("Default constructor")
Friend Sub CreInit(ByVal connString As String, ByVal factory As IDbConnectionFactory, ByVal baseCommand As IDbCommandBase)
Attribute CreInit.VB_Description = "Default constructor"
    Errors.GuardEmptyString connString
    Errors.GuardNullReference factory
    Errors.GuardNullReference baseCommand

    this.ConnectionString = connString
    Set this.ConnectionFactory = factory
    Set this.Base = baseCommand
End Sub


Private Function IDbCommand_Execute(ByVal sql As String, ParamArray args() As Variant) As ADODB.Recordset
    Dim localArgs() As Variant
    localArgs = args 'cannot pass ParamArray array directly to ByRef array argument
    
    With this.ConnectionFactory.CreateInstance(this.ConnectionString)
        Dim cmd As ADODB.Command
        Set cmd = this.Base.CreateCommand(.Self, adCmdText, sql, localArgs)
        
        Dim result As ADODB.Recordset
        Set result = this.Base.GetDisconnectedRecordset(cmd)
        
        Set IDbCommand_Execute = result
    End With
End Function


Private Sub IDbCommand_ExecuteNonQuery(ByVal sql As String, ParamArray args() As Variant)
    Dim localArgs() As Variant
    localArgs = args 'cannot pass ParamArray array directly to ByRef array argument
    
    With this.ConnectionFactory.CreateInstance(this.ConnectionString)
        Dim cmd As ADODB.Command
        Set cmd = this.Base.CreateCommand(.Self, adCmdText, sql, localArgs)
        
        '@Ignore FunctionReturnValueDiscarded
        cmd.Execute
    End With
End Sub


Private Function IDbCommand_ExecuteWithParameters(ByVal sql As String, ParamArray params() As Variant) As ADODB.Recordset
    Dim localParams() As Variant
    localParams = params

    With this.ConnectionFactory.CreateInstance(this.ConnectionString)
        Dim emptyArgs() As Variant
        
        Dim cmd As ADODB.Command
        Set cmd = this.Base.CreateCommand(.Self, adCmdText, sql, emptyArgs)
        
        Dim i As Long
        For i = LBound(localParams) To UBound(localParams)
            cmd.Parameters.Append localParams(i)
        Next
        
        Dim result As ADODB.Recordset
        Set result = this.Base.GetDisconnectedRecordset(cmd)
        
        Set IDbCommand_ExecuteWithParameters = result
    End With
End Function


Private Function IDbCommand_GetSingleValue(ByVal sql As String, ParamArray args() As Variant) As Variant
    'cannot pass ParamArray array directly to ByRef array argument
    Dim localArgs() As Variant
    localArgs = args
    
    With this.ConnectionFactory.CreateInstance(this.ConnectionString)
        IDbCommand_GetSingleValue = this.Base.GetSingleValue(.Self, sql, localArgs)
    End With
End Function
