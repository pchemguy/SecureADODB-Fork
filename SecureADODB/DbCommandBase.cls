VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbCommandBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object intended to be used as a private instance field in a class implementing the IDbCommand interface (composition in place of inheritance)."
'@Folder "SecureADODB.DbCommand.Abstract"
'@ModuleDescription("An object intended to be used as a private instance field in a class implementing the IDbCommand interface (composition in place of inheritance).")
'@PredeclaredId
'@IgnoreModule UseMeaningfulName, IllegalAnnotation
'@Exposed
Option Explicit

Implements IDbCommandBase

Private Type TCommand
    ParameterProvider As IParameterProvider
End Type
Private this As TCommand


'@Inject: just an idea.. see #https://github.com/rubberduck-vba/Rubberduck/issues/5463
Public Function Create(ByVal provider As IParameterProvider) As IDbCommandBase
    Errors.GuardNonDefaultInstance Me, DbCommandBase
    
    Dim result As DbCommandBase
    Set result = New DbCommandBase
    result.CreInit provider
    Set Create = result
End Function


'@Description("Default constructor")
Friend Sub CreInit(ByVal provider As IParameterProvider)
Attribute CreInit.VB_Description = "Default constructor"
    Errors.GuardNullReference provider
    Set this.ParameterProvider = provider
End Sub


Private Function CreateCommand(ByVal db As IDbConnection, ByVal CommandType As ADODB.CommandTypeEnum, ByVal sql As String, ByRef args() As Variant) As ADODB.Command
    Errors.GuardNullReference db
    Errors.GuardExpression db.State <> adStateOpen, message:="Connection is not open."
    Errors.GuardEmptyString sql
    Errors.GuardExpression Not ValidateOrdinalArguments(sql, args), message:="Arguments supplied are inconsistent with the provided command string parameters."
    
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = db.AdoConnection
    cmd.CommandType = CommandType
    cmd.CommandText = sql
    
    On Error GoTo CleanFail
    Dim arg As ADODB.Parameter
    For Each arg In this.ParameterProvider.FromValues(args)
        cmd.Parameters.Append arg
    Next
    
CleanExit:
    Set CreateCommand = cmd
    Exit Function
CleanFail:
    Resume CleanExit
End Function


Private Function GetDisconnectedRecordset(ByVal cmd As ADODB.Command) As ADODB.Recordset
    Errors.GuardNullReference cmd
    Errors.GuardNullReference cmd.ActiveConnection
    
    Dim result As ADODB.Recordset
    Set result = New ADODB.Recordset
    
    result.CursorLocation = adUseClient
    result.Open source:=cmd, CursorType:=adOpenStatic
    
    Set result.ActiveConnection = Nothing
    Set GetDisconnectedRecordset = result
End Function


Private Function GetSingleValue(ByVal db As IDbConnection, ByVal CommandType As ADODB.CommandTypeEnum, ByVal sql As String, ByRef args() As Variant) As Variant
    Errors.GuardEmptyString sql
    
    Dim cmd As ADODB.Command
    Set cmd = CreateCommand(db, CommandType, sql, args)
    
    Dim results As ADODB.Recordset
    Set results = GetDisconnectedRecordset(cmd)
    
    GetSingleValue = results.Fields.Item(0).value
End Function


Public Function ValidateOrdinalArguments(ByVal sql As String, ByRef args() As Variant) As Boolean
    On Error GoTo CleanFail
    Dim result As Boolean
    
    Dim expected As Long
    expected = Len(sql) - Len(Replace(sql, "?", vbNullString))
    
    Dim actual As Long
    On Error GoTo CleanFail 'if there are no args, LBound/UBound are both out of bounds
    actual = UBound(args) + (1 - LBound(args))
    
CleanExit:
    result = (expected = actual)
    ValidateOrdinalArguments = result
    Exit Function
CleanFail:
    actual = 0
    Resume CleanExit
End Function


Private Function IDbCommandBase_CreateCommand(ByVal db As IDbConnection, ByVal CommandType As ADODB.CommandTypeEnum, ByVal sql As String, ByRef args() As Variant) As ADODB.Command
    Set IDbCommandBase_CreateCommand = CreateCommand(db, CommandType, sql, args)
End Function

Private Function IDbCommandBase_GetDisconnectedRecordset(ByVal cmd As ADODB.Command) As ADODB.Recordset
    Set IDbCommandBase_GetDisconnectedRecordset = GetDisconnectedRecordset(cmd)
End Function

Private Function IDbCommandBase_GetSingleValue(ByVal db As IDbConnection, ByVal CommandType As ADODB.CommandTypeEnum, ByVal sql As String, ByRef args() As Variant) As Variant
    IDbCommandBase_GetSingleValue = GetSingleValue(db, CommandType, sql, args)
End Function
