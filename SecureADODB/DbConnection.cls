VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A thin wrapper around an ADODB connection."
'@Folder "SecureADODB.DbConnection"
'@ModuleDescription("A thin wrapper around an ADODB connection.")
'@PredeclaredId
'@IgnoreModule AssignmentNotUsed, IllegalAnnotation
'@Exposed
Option Explicit

Implements IDbConnection

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents AdoConnection As ADODB.Connection
Attribute AdoConnection.VB_VarHelpID = -1
Private Type TDbConnection
    LogController As ILogger
    '''' If  particular backend does not support transaction, "begin" method raises an error, and the corresponding handler sets this flag
    TransactionsDisabled As Boolean
    HasActiveTransaction As Boolean
    '''' Map for ADODB.ObjectStateEnum
    StateMap As Variant
End Type
Private this As TDbConnection


'@Inject: just an idea.. see #https://github.com/rubberduck-vba/Rubberduck/issues/5463
Public Function Create(ByVal connString As String, Optional ByVal LogController As ILogger = Nothing) As IDbConnection
    Errors.GuardNonDefaultInstance Me, DbConnection
    
    Dim result As DbConnection
    Set result = New DbConnection
    result.CreInit connString, LogController
    Set Create = result
End Function


'@Description("Default constructor")
Friend Sub CreInit(ByVal connString As String, Optional ByVal LogController As ILogger = Nothing)
Attribute CreInit.VB_Description = "Default constructor"
    Errors.GuardEmptyString connString
    
    Set this.LogController = LogController
    this.TransactionsDisabled = False
    AdoConnection.Open connString
End Sub


Private Sub Class_Initialize()
    Set AdoConnection = New ADODB.Connection
    '@Ignore EmptyStringLiteral
    this.StateMap = Array("adStateClosed", "adStateOpen", "adStateConnecting", "", "adStateExecuting", "", "", "", "adStateFetching")
End Sub


Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    If this.HasActiveTransaction Then AdoConnection.RollbackTrans
    AdoConnection.Close
    Set AdoConnection = Nothing
    
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    LogAdoErrors
    Resume Next
End Sub


Private Sub LogAdoErrors()
    If this.LogController Is Nothing Then Exit Sub
    
    Dim AdoError As ADODB.Error
    Dim message As String
    message = "ADODB.Connection State: " + CStr(this.StateMap(AdoConnection.State))
    this.LogController.Log message
    For Each AdoError In AdoConnection.Errors
        With AdoError
            message = "ADODB.Error #" + CStr(.number) + ", " + _
                      "SQLState: " + CStr(.SqlState) + ", " + _
                      "Source: " + CStr(.source) + ", " + _
                      "Description: " + CStr(.description)
        End With
        this.LogController.Log message
    Next AdoError
    AdoConnection.Errors.Clear
        
    With VBA.Err
        If .number > 0 And Left(.source, 6) = "ADODB." Then
            message = "VBA.Err #" + CStr(.number) + ", " + _
                      "Source: " + CStr(.source) + ", " + _
                      "Description: " + CStr(.description)
        End If
    End With
    this.LogController.Log message
    VBA.Err.Clear
End Sub


'================================================================================'
'============================ IDbConnection INTERFACE ==========================='
'================================================================================'

Private Property Get IDbConnection_Self() As IDbConnection
    Set IDbConnection_Self = Me
End Property

Private Property Get IDbConnection_AdoConnection() As ADODB.Connection
    Set IDbConnection_AdoConnection = AdoConnection
End Property

Private Property Get IDbConnection_State() As ADODB.ObjectStateEnum
    IDbConnection_State = AdoConnection.State
End Property

Private Sub IDbConnection_BeginTransaction()
    Guard.inTransaction this.HasActiveTransaction
    
    On Error GoTo ErrorHandler
    '@Ignore VariableNotUsed
    Dim result As Long
    '''' If  particular backend does not support transaction, the next line raises an error, and the corresponding handler sets this.TransactionsDisabled
    result = AdoConnection.BeginTrans
    If Not this.TransactionsDisabled Then this.HasActiveTransaction = True
    LogAdoErrors
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    If VBA.Err.number = 3251 Then
        this.TransactionsDisabled = True
        Resume Next
    Else
        Err.Raise Err.number, Err.source, Err.description, Err.HelpFile, Err.HelpContext
    End If
End Sub

Private Sub IDbConnection_CommitTransaction()
    If this.TransactionsDisabled Then Exit Sub
    Guard.NotInTransaction this.HasActiveTransaction
    
    AdoConnection.CommitTrans
    this.HasActiveTransaction = False
    LogAdoErrors
End Sub

Private Sub IDbConnection_RollbackTransaction()
    If this.TransactionsDisabled Then Exit Sub
    Guard.NotInTransaction this.HasActiveTransaction
    
    AdoConnection.RollbackTrans
    this.HasActiveTransaction = False
    LogAdoErrors
End Sub


'================================================================================'
'========================= AdoConnection EVENT HANDLERS ========================='
'================================================================================'

Private Sub AdoConnection_BeginTransComplete(ByVal TransactionLevel As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Begin transaction completed. "
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.number, pError.description
End Sub

Private Sub AdoConnection_CommitTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Commit transaction completed. "
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.number, pError.description
End Sub

Private Sub AdoConnection_RollbackTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Rollback transaction completed."
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.number, pError.description
End Sub

Private Sub AdoConnection_WillConnect(ConnectionString As String, UserID As String, Password As String, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Begin connect..."
End Sub

Private Sub AdoConnection_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Connect completed. Status: " & adStatus
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.number, pError.description
End Sub

Private Sub AdoConnection_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Disconnect completed. Status: " & adStatus
End Sub

Private Sub AdoConnection_WillExecute(source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Debug.Print "Begin execute..."
End Sub

Private Sub AdoConnection_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Debug.Print "Execute completed, " & RecordsAffected & " record(s) affected."
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.number, pError.description
End Sub

Private Sub AdoConnection_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    If pError.number = 0 Then
        Debug.Print "Info: " & pError.description
    Else
        Debug.Print "Error: " & pError.number, pError.description
    End If
End Sub
