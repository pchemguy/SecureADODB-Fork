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
'@Folder "SecureADODBmod.DbManager"
'@ModuleDescription("An object that encapsulates a database transaction.")
'@PredeclaredId
'@IgnoreModule ProcedureNotUsed
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
'@Ignore MoveFieldCloserToUsage
Private connConfig As TConnectionConfig

Private Type TDbManager
    Connection As IDbConnection
    CommandFactory As IDbCommandFactory
    Committed As Boolean
    RolledBack As Boolean
End Type
Private this As TDbManager


'@Description("Returns class reference")
Public Property Get Class() As DbManager
Attribute Class.VB_Description = "Returns class reference"
    Set Class = DbManager
End Property


'@Description("Build connection config")
Public Function BuildConnectionString( _
                 ByVal Flavor As String, _
        Optional ByVal Path As String = vbNullString, _
        Optional ByVal Name As String = vbNullString, _
        Optional ByVal Options As String = vbNullString _
        ) As String
Attribute BuildConnectionString.VB_Description = "Build connection config"
        
    Errors.GuardEmptyString Flavor
        
    If Path = vbNullString Then
        connConfig.Path = ThisWorkbook.Path
    Else
        connConfig.Path = Path
    End If
    Select Case LCase(Flavor)
        Case "sqlite"
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
            End With
        Case "csv"
            connConfig.Driver = "{Microsoft Text Driver (*.txt; *.csv)}"
            connConfig.Database = connConfig.Path
            If Options <> vbNullString Then
                connConfig.Options = Options
            End If
        Case Else
            Dim errorDetails As TError
            With errorDetails
                .number = ErrNo.FeatureNotAvailableErr
                .Name = "FeatureNotAvailable"
                .source = "DbManager.BuildConnectionString"
                .message = "Config assistance for database " & Flavor & " is not available. Use raw connection string."
                .description = .message
            End With
            RaiseError errorDetails
    End Select
    With connConfig
        .connString = "Driver=" + .Driver + ";" + _
                      "Database=" + .Database + ";" + _
                      .Options
    End With
    BuildConnectionString = connConfig.connString
End Function


'@Description("Creates a new manager using connection parameters.")
Public Function FromConnectionParameters( _
                 ByVal Flavor As String, _
        Optional ByVal Path As String = vbNullString, _
        Optional ByVal Name As String = vbNullString, _
        Optional ByVal Options As String = vbNullString _
        ) As IDbManager
Attribute FromConnectionParameters.VB_Description = "Creates a new manager using connection parameters."

    Set FromConnectionParameters = DbManager.FromConnectionString( _
            BuildConnectionString(Flavor, Path, Name, Options))
End Function


'@Ignore ProcedureNotUsed
'@Description("Creates a new manager using default configurations.")
Public Function FromConnectionString(ByVal connString As String) As IDbManager
Attribute FromConnectionString.VB_Description = "Creates a new manager using default configurations."
    Errors.GuardNonDefaultInstance Me, DbManager
    
    '@Ignore UseMeaningfulName
    Dim db As IDbConnection
    Set db = DbConnection.Create(connString)
    
    Dim provider As IParameterProvider
    Set provider = AdoParameterProvider.Create(AdoTypeMappings.Default)
    
    Dim baseCommand As IDbCommandBase
    Set baseCommand = DbCommandBase.Create(provider)
    
    Dim factory As IDbCommandFactory
    Set factory = DefaultDbCommandFactory.Create(baseCommand)
    
    Set FromConnectionString = DbManager.Create(db, factory)
End Function


'@Ignore UseMeaningfulName
Public Function Create(ByVal db As IDbConnection, ByVal factory As IDbCommandFactory) As IDbManager
    Guard.NonDefaultInstance Me
    
    Dim result As DbManager
    Set result = New DbManager
    result.CreInit db, factory
    Set Create = result
End Function


'@Ignore UseMeaningfulName
'@Description("Default constructor")
Friend Sub CreInit(ByVal db As IDbConnection, ByVal factory As IDbCommandFactory)
Attribute CreInit.VB_Description = "Default constructor"
    Guard.NullReference factory
    Guard.NullReference db
    Errors.GuardExpression db.State <> adStateOpen, message:="Connection should be open."

    Set this.Connection = db
    Set this.CommandFactory = factory
    this.Connection.BeginTransaction
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
    Set IDbManager_Command = this.CommandFactory.CreateInstance(this.Connection)
End Function


Private Sub IDbManager_Commit()
    Errors.GuardExpression this.Committed, message:="Transaction is already committed."
    Errors.GuardExpression this.RolledBack, message:="Transaction was rolled back."
    On Error Resume Next ' not all providers support transactions
    this.Connection.CommitTransaction
    this.Committed = True
    On Error GoTo 0
End Sub


Private Sub IDbManager_Rollback()
    Errors.GuardExpression this.Committed, message:="Transaction is already committed."
    
    On Error Resume Next ' not all providers support transactions
    this.Connection.RollbackTransaction
    this.RolledBack = True
    On Error GoTo 0
End Sub
