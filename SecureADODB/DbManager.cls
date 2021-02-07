VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An object that encapsulates a database transaction."
'@Folder "SecureADODB.DbManager"
'@ModuleDescription("An object that encapsulates a database transaction.")
'@PredeclaredId
'@IgnoreModule UseMeaningfulName, ProcedureNotUsed
Option Explicit

Implements IDbManager

Private Type TConnectionConfig
    Path As String
    Name As String
    Database As String
    Driver As String
    Options As String
    connString As String
End Type

Private Type TDbManager
    Connection As IDbConnection
    CommandFactory As IDbCommandFactory
    UseTransactions As Boolean
    UseDisconnectedRecordset As Boolean
    Committed As Boolean
    RolledBack As Boolean
End Type
Private this As TDbManager


'@Description("Returns class reference")
Public Property Get Class() As DbManager
Attribute Class.VB_Description = "Returns class reference"
    Set Class = DbManager
End Property


' Basic connection string construction facility
'
' The first argument specifies either the type of the target database or, for reasons of
' flexibility can take a completed connection string. Select Case implements the core
' functionality. If the first argument matches any of the Case block, the corresponding
' configuration will be applied and a connection string is built and returned. Otherwise,
' if the first argument contains "Driver=" substring, it is assumed to be a completed
' connection string, which is returned as is.
'
'@Description("Build connection config")
Public Function BuildConnectionString( _
                 ByVal TypeOrConnString As String, _
        Optional ByVal Path As String = vbNullString, _
        Optional ByVal Name As String = vbNullString, _
        Optional ByVal Options As String = vbNullString _
        ) As String
Attribute BuildConnectionString.VB_Description = "Build connection config"
        
    Guard.EmptyString TypeOrConnString
    Dim connConfig As TConnectionConfig
        
    If Path = vbNullString Then
        connConfig.Path = ThisWorkbook.Path
    Else
        connConfig.Path = Path
    End If
    Select Case LCase(TypeOrConnString)
        Case "sqlite"                                                           ' ========== SQLite db ========== '
            With connConfig
                .Driver = "{SQLite3 ODBC Driver}"
                If Name = vbNullString Then
                    Dim fso As Scripting.FileSystemObject
                    Set fso = New Scripting.FileSystemObject
                    .Name = fso.GetBaseName(ThisWorkbook.Name) & ".db"
                Else
                    .Name = Name
                End If
                If Options = vbNullString Then
                    .Options = "SyncPragma=NORMAL;LongNames=True;NoCreat=True;FKSupport=True;OEMCP=True;"
                Else
                    .Options = Options
                End If
                .Database = .Path & Application.PathSeparator & .Name
                .connString = "Driver=" + .Driver + ";" + _
                              "Database=" + .Database + ";" + _
                              .Options
            End With
        Case "csv"                                                              ' ============ CSV db ============ '
            With connConfig
                .Driver = "{Microsoft Text Driver (*.txt; *.csv)}"
                .Database = .Path
                If Options <> vbNullString Then
                    .Options = Options
                End If
                .connString = "Driver=" + .Driver + ";" + _
                              "DefaultDir=" + .Database + ";" + _
                              .Options
            End With
        Case Else
            If InStr(1, TypeOrConnString, "Driver=", vbTextCompare) > 0 Then    ' ============ Raw Connection String ============ '
                connConfig.connString = TypeOrConnString
            Else
                connConfig.connString = vbNullString
            End If
    End Select
    BuildConnectionString = connConfig.connString
End Function


'@Description("Creates a new manager using connection parameters.")
Public Function FromConnectionParameters( _
                 ByVal TypeOrConnString As String, _
        Optional ByVal Path As String = vbNullString, _
        Optional ByVal Name As String = vbNullString, _
        Optional ByVal Options As String = vbNullString, _
        Optional ByVal UseTransactions As Boolean = True, _
        Optional ByVal UseDisconnectedRecordset As Boolean = False _
        ) As IDbManager
Attribute FromConnectionParameters.VB_Description = "Creates a new manager using connection parameters."

    Guard.NonDefaultInstance Me
    
    Dim connString As String
    connString = BuildConnectionString(TypeOrConnString, Path, Name, Options)
    
    Dim db As IDbConnection
    Set db = DbConnection.Create(connString)
    
    Dim provider As IParameterProvider
    Set provider = AdoParameterProvider.Create(AdoTypeMappings.Default)
    
    Dim baseCommand As IDbCommandBase
    Set baseCommand = DbCommandBase.Create(provider)
    
    Dim factory As IDbCommandFactory
    Set factory = DbCommandFactory.Create(baseCommand)
    
    Set FromConnectionParameters = DbManager.Create(db, factory, UseTransactions, UseDisconnectedRecordset)
End Function


'@Ignore UseMeaningfulName
Public Function Create(ByVal db As IDbConnection, _
                       ByVal factory As IDbCommandFactory, _
              Optional ByVal UseTransactions As Boolean = True, _
              Optional ByVal UseDisconnectedRecordset As Boolean = False) As IDbManager
              
    Guard.NonDefaultInstance Me
    
    Dim result As DbManager
    Set result = New DbManager
    result.CreInit db, factory, UseTransactions, UseDisconnectedRecordset
    Set Create = result
End Function


'@Ignore UseMeaningfulName
'@Description("Default constructor")
Friend Sub CreInit(ByVal db As IDbConnection, _
                       ByVal factory As IDbCommandFactory, _
              Optional ByVal UseTransactions As Boolean = True, _
              Optional ByVal UseDisconnectedRecordset As Boolean = False)
Attribute CreInit.VB_Description = "Default constructor"
              
    Guard.NullReference factory
    Guard.NullReference db
    Errors.GuardExpression db.State <> adStateOpen, message:="Connection should be open."

    Set this.Connection = db
    Set this.CommandFactory = factory
    this.UseTransactions = UseTransactions
    this.UseDisconnectedRecordset = UseDisconnectedRecordset
    If this.UseTransactions Then this.Connection.BeginTransaction
End Sub


Private Sub Class_Terminate()
    On Error Resume Next
    If Not this.Committed Then this.Connection.RollbackTransaction
    On Error GoTo 0
End Sub


'@Description("Returns class reference")
Public Property Get IDbManager_Class() As DbManager
Attribute IDbManager_Class.VB_Description = "Returns class reference"
    Set IDbManager_Class = DbManager
End Property


Private Function IDbManager_Command() As IDbCommand
    Set IDbManager_Command = this.CommandFactory.CreateInstance(this.Connection, this.UseDisconnectedRecordset)
End Function


Private Sub IDbManager_Commit()
    Errors.GuardExpression (Not this.UseTransactions), message:="Transactions are disabled by the caller."
    Errors.GuardExpression this.Committed, message:="Transaction is already committed."
    Errors.GuardExpression this.RolledBack, message:="Transaction was rolled back."
    On Error Resume Next ' not all providers support transactions
    this.Connection.CommitTransaction
    this.Committed = True
    On Error GoTo 0
End Sub


Private Sub IDbManager_Rollback()
    Errors.GuardExpression (Not this.UseTransactions), message:="Transactions are disabled by the caller."
    Errors.GuardExpression this.Committed, message:="Transaction is already committed."
    
    On Error Resume Next ' not all providers support transactions
    this.Connection.RollbackTransaction
    this.RolledBack = True
    On Error GoTo 0
End Sub
