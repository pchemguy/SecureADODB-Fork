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
'@IgnoreModule UseMeaningfulName, IllegalAnnotation
'@Exposed
Option Explicit

Implements IDbConnection

'@Ignore MissingMemberAnnotation
Private WithEvents AdoConnection As ADODB.Connection
Attribute AdoConnection.VB_VarHelpID = -1
Private HasActiveTransaction As Boolean


'@Inject: just an idea.. see #https://github.com/rubberduck-vba/Rubberduck/issues/5463
Public Function Create(ByVal connString As String) As IDbConnection
    Errors.GuardNonDefaultInstance Me, DbConnection
    
    Dim result As DbConnection
    Set result = New DbConnection
    result.CreInit connString
    Set Create = result
End Function


'@Description("Default constructor")
Friend Sub CreInit(ByVal connString As String)
Attribute CreInit.VB_Description = "Default constructor"
    Errors.GuardEmptyString connString
    AdoConnection.Open connString
End Sub


Private Sub Class_Initialize()
    Set AdoConnection = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    AdoConnection.Close
    On Error GoTo 0
    Set AdoConnection = Nothing
End Sub


Private Sub AdoConnection_BeginTransComplete(ByVal TransactionLevel As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Begin transaction completed. "
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.number, pError.description
End Sub

Private Sub AdoConnection_CommitTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Commit transaction completed. "
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.number, pError.description
End Sub

Private Sub AdoConnection_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Connect completed. Status: " & adStatus
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.number, pError.description
End Sub

Private Sub AdoConnection_Disconnect(adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Disconnect completed. Status: " & adStatus
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

Private Sub AdoConnection_RollbackTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Rollback transaction completed."
    If Not pError Is Nothing Then Debug.Print "Error: " & pError.number, pError.description
End Sub

Private Sub AdoConnection_WillConnect(ConnectionString As String, UserID As String, Password As String, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print "Begin connect..."
End Sub

Private Sub AdoConnection_WillExecute(source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Debug.Print "Begin execute..."
End Sub


Private Property Get IDbConnection_Self() As IDbConnection
    Set IDbConnection_Self = Me
End Property

Private Property Get IDbConnection_AdoConnection() As ADODB.Connection
    Set IDbConnection_AdoConnection = AdoConnection
End Property

Private Sub IDbConnection_BeginTransaction()
    On Error Resume Next 'not all providers support transactions
    '@Ignore VariableNotUsed
    Dim result As Long
    result = AdoConnection.BeginTrans
    On Error GoTo 0
    HasActiveTransaction = True
End Sub

Private Sub IDbConnection_CommitTransaction()
    On Error Resume Next
    AdoConnection.CommitTrans
    On Error GoTo 0
    HasActiveTransaction = False
End Sub

Private Sub IDbConnection_RollbackTransaction()
    If HasActiveTransaction Then AdoConnection.RollbackTrans
    HasActiveTransaction = False
End Sub

Private Property Get IDbConnection_State() As ADODB.ObjectStateEnum
    IDbConnection_State = AdoConnection.State
End Property
