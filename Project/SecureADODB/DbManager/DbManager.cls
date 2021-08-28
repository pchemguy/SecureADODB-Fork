VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object that encapsulates a database transaction."
'@Folder "SecureADODB.DbManager"
'@ModuleDescription("An object that encapsulates a database transaction.")
'@PredeclaredId
'@Exposed
'@IgnoreModule ProcedureNotUsed
Option Explicit

Implements IDbManager

Private Type TDbManager
    Path As String
    Name As String
    Database As String
    Driver As String
    Options As String
    ConnectionString As String
    Connection As IDbConnection
    CommandFactory As IDbCommandFactory
    UseTransactions As Boolean
    LogController As ILogger
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
        
    Guard.ConnectionString TypeOrConnString
    Dim connConfig As TDbManager
        
    If Path = vbNullString Then
        connConfig.Path = ThisWorkbook.Path
    Else
        connConfig.Path = Path
    End If
    Select Case LCase$(TypeOrConnString)
        Case "sqlite"                                                           ' ========== SQLite db ========== '
            With connConfig
                .Driver = "SQLite3 ODBC Driver"
                If Name = vbNullString Then
                    Dim fso As Scripting.FileSystemObject
                    Set fso = New Scripting.FileSystemObject
                    .Name = fso.GetBaseName(ThisWorkbook.Name) & ".db"
                Else
                    .Name = Name
                End If
                If Options = vbNullString Then
                    .Options = "SyncPragma=NORMAL;FKSupport=True;"
                Else
                    .Options = Options
                End If
                .Database = .Path & Application.PathSeparator & .Name
                .ConnectionString = "Driver=" + .Driver + ";" + _
                              "Database=" + .Database + ";" + _
                              .Options
            End With
        Case "csv"                                                              ' ============ CSV db ============ '
            With connConfig
                #If Win64 Then
                    .Driver = "Microsoft Access Text Driver (*.txt, *.csv)"
                #Else
                    .Driver = "{Microsoft Text Driver (*.txt; *.csv)}"
                #End If
                .Database = .Path
                If Options <> vbNullString Then
                    .Options = Options
                End If
                .ConnectionString = "Driver=" + .Driver + ";" + _
                              "DefaultDir=" + .Database + ";" + _
                              .Options
            End With
        Case Else
            If InStr(1, TypeOrConnString, "Driver=", vbTextCompare) > 0 Or _
               InStr(1, TypeOrConnString, "DSN=", vbTextCompare) > 0 Then    ' ============ Raw Connection String ============ '
                connConfig.ConnectionString = TypeOrConnString
            Else
                Guard.ConnectionString vbNullString
            End If
    End Select
    BuildConnectionString = connConfig.ConnectionString
End Function


'@Description("Creates a new manager using connection parameters.")
Public Function FromConnectionParameters( _
                 ByVal TypeOrConnString As String, _
        Optional ByVal Path As String = vbNullString, _
        Optional ByVal Name As String = vbNullString, _
        Optional ByVal Options As String = vbNullString, _
        Optional ByVal UseTransactions As Boolean = True, _
        Optional ByVal LoggerType As LoggerTypeEnum = LoggerTypeEnum.logGlobal _
        ) As IDbManager
Attribute FromConnectionParameters.VB_Description = "Creates a new manager using connection parameters."

    Guard.NonDefaultInstance Me
    
    Dim connString As String
    connString = BuildConnectionString(TypeOrConnString, Path, Name, Options)
    
    Dim LogController As ILogger
    Select Case LoggerType
        Case LoggerTypeEnum.logDisabled
            Set LogController = Nothing
        Case LoggerTypeEnum.logGlobal
            Set LogController = Logger
        Case LoggerTypeEnum.logPrivate
            Set LogController = Logger.Create
    End Select
    
    Dim provider As IParameterProvider
    Set provider = AdoParameterProvider.Create(AdoTypeMappings.Default)
    
    Dim baseCommand As IDbCommandBase
    Set baseCommand = DbCommandBase.Create(provider)
    
    Dim factory As IDbCommandFactory
    Set factory = DbCommandFactory.Create(baseCommand)
    
    Dim db As IDbConnection
    Set db = DbConnection.Create(connString, LogController)
    
    Set FromConnectionParameters = DbManager.Create(db, factory, UseTransactions, LogController)
End Function


Public Function Create(ByVal db As IDbConnection, _
                       ByVal factory As IDbCommandFactory, _
              Optional ByVal UseTransactions As Boolean = True, _
              Optional ByVal LogController As ILogger = Nothing) As IDbManager
              
    Guard.NonDefaultInstance Me
    
    Dim Result As DbManager
    Set Result = New DbManager
    Result.Init db, factory, UseTransactions, LogController
    Set Create = Result
End Function


'@Description("Default constructor")
Friend Sub Init(ByVal db As IDbConnection, _
                ByVal factory As IDbCommandFactory, _
       Optional ByVal UseTransactions As Boolean = True, _
       Optional ByVal LogController As ILogger = Nothing)
Attribute Init.VB_Description = "Default constructor"
              
    Guard.NullReference factory
    Guard.NullReference db
    Guard.Expression db.State = adStateOpen, Source:="DbManager", Message:="Connection should be open."

    Set this.LogController = LogController
    Set this.Connection = db
    Set this.CommandFactory = factory
    this.UseTransactions = UseTransactions
    If this.UseTransactions Then this.Connection.BeginTransaction
End Sub


'@Description("Returns class reference")
Public Property Get IDbManager_Class() As DbManager
Attribute IDbManager_Class.VB_Description = "Returns class reference"
    Set IDbManager_Class = DbManager
End Property


Private Property Get IDbManager_LogController() As ILogger
    Set IDbManager_LogController = this.LogController
End Property


Private Property Get IDbManager_Connection() As IDbConnection
    Set IDbManager_Connection = this.Connection
End Property


Private Function IDbManager_Command() As IDbCommand
    Set IDbManager_Command = this.CommandFactory.CreateInstance(this.Connection)
End Function


Private Function IDbManager_Recordset(Optional ByVal Scalar As Boolean = False, _
                                      Optional ByVal Disconnected As Boolean = True, _
                                      Optional ByVal CacheSize As Long = 10, _
                                      Optional ByVal CursorType As ADODB.CursorTypeEnum = -1, _
                                      Optional ByVal AsyncMode As Boolean = False, _
                                      Optional ByVal AsyncOption As ADODB.ExecuteOptionEnum = 0) As IDbRecordset
              
    Dim cmd As IDbCommand
    Set cmd = this.CommandFactory.CreateInstance(this.Connection)
    Set IDbManager_Recordset = DbRecordset.Create( _
        cmd, Scalar, Disconnected, CacheSize, CursorType, AsyncMode, AsyncOption)
End Function


Private Sub IDbManager_Commit()
    Guard.Expression this.UseTransactions, Source:="DbManager", Message:="Transactions are disabled by the caller."
    this.Connection.CommitTransaction
End Sub


Private Sub IDbManager_Rollback()
    Guard.Expression this.UseTransactions, Source:="DbManager", Message:="Transactions are disabled by the caller."
    this.Connection.RollbackTransaction
End Sub
