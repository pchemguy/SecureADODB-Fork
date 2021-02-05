VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitOfWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object that encapsulates a database transaction."
'@Folder "SecureADODB.UnitOfWork"
'@ModuleDescription("An object that encapsulates a database transaction.")
'@PredeclaredId
'@IgnoreModule UseMeaningfulName, IllegalAnnotation
'@Exposed
Option Explicit

Implements IUnitOfWork

Private Type TUnitOfWork
    Connection As IDbConnection
    CommandFactory As IDbCommandFactory
    Committed As Boolean
    RolledBack As Boolean
End Type
Private this As TUnitOfWork


'@Description("Creates a new unit of work using default configurations.")
'@Ignore ProcedureNotUsed
Public Function FromConnectionString(ByVal connString As String) As IUnitOfWork
Attribute FromConnectionString.VB_Description = "Creates a new unit of work using default configurations."
    Dim db As IDbConnection
    Set db = DbConnection.Create(connString)
    
    Dim provider As IParameterProvider
    Set provider = AdoParameterProvider.Create(AdoTypeMappings.Default)
    
    Dim baseCommand As IDbCommandBase
    Set baseCommand = DbCommandBase.Create(provider)
    
    Dim factory As IDbCommandFactory
    Set factory = DefaultDbCommandFactory.Create(baseCommand)
    
    Set FromConnectionString = UnitOfWork.Create(db, factory)
End Function


'@Inject: just an idea.. see #https://github.com/rubberduck-vba/Rubberduck/issues/5463
Public Function Create(ByVal db As IDbConnection, ByVal factory As IDbCommandFactory) As IUnitOfWork
    Errors.GuardNonDefaultInstance Me, UnitOfWork
    
    Dim result As UnitOfWork
    Set result = New UnitOfWork
    result.CreInit db, factory
    Set Create = result
End Function


'@Description("Default constructor")
Friend Sub CreInit(ByVal db As IDbConnection, ByVal factory As IDbCommandFactory)
Attribute CreInit.VB_Description = "Default constructor"
    Errors.GuardNullReference factory
    Errors.GuardNullReference db
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


Private Sub IUnitOfWork_Commit()
    Errors.GuardExpression this.Committed, message:="Transaction is already committed."
    Errors.GuardExpression this.RolledBack, message:="Transaction was rolled back."
    On Error Resume Next ' not all providers support transactions
    this.Connection.CommitTransaction
    this.Committed = True
    On Error GoTo 0
End Sub


Private Function IUnitOfWork_Command() As IDbCommand
    Set IUnitOfWork_Command = this.CommandFactory.CreateInstance(this.Connection)
End Function


Private Sub IUnitOfWork_Rollback()
    Errors.GuardExpression this.Committed, message:="Transaction is already committed."
    On Error Resume Next ' not all providers support transactions
    this.Connection.RollbackTransaction
    this.RolledBack = True
    On Error GoTo 0
End Sub
