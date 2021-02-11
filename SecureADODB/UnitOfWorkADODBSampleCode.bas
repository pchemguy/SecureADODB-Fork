Attribute VB_Name = "UnitOfWorkADODBSampleCode"
'@Folder("-- DraftsTemplatesSnippets --")
'@IgnoreModule EmptyModule, VariableNotUsed, ProcedureNotUsed
Option Explicit


    
'Private Sub DbManager()
'    Dim connString As String
'    connString = ConnectionStringObject.ADOConnectionString
'
'    Dim SQLQuery As String
'    SQLQuery = "SELECT * FROM categories WHERE category_id <= 3 AND section = 'machinery'"
'
''    Dim SQLRecordset As ADODB.Recordset
''    Set SQLRecordset = UnitOfWork.FromConnectionString(connString).Command.Execute(SQLQuery)
'
'    With DbManager.FromConnectionString(connString)
'        'connection is open, a transaction is initiated.
'
'        'IDbCommand.Execute returns a disconnected ADODB.Recordset:
'        Dim results As ADODB.Recordset
'        'simply use '?' ordinal parameters in the command string, and then provide a value for each '?' in the SQL:
'        Set results = .Command.Execute(SQLQuery)
'
'        'we are in a transaction, so we need to commit the changes - lest we lose them:
'        .Commit '<~ make sure to only commit AT MOST ONCE per transaction.
'
'        Dim rows As Variant
'        rows = results.GetRows
'    End With 'transaction is rolled back if not committed, connection is closed.
'End Sub


'Private Sub AutoDbCommandTest()
'    Dim ConnectionStringObject As SqliteConnectionString
'    Set ConnectionStringObject = SqliteConnectionString.Create(ThisWorkbook.Path, "SecureADODB.db")
'    Dim connString As String
'    connString = ConnectionStringObject.ADOConnectionString
'
'    Dim SQLQuery As String
'    SQLQuery = "SELECT * FROM categories WHERE category_id <= 3 AND section = 'machinery'"
'
'    Dim mappings As ITypeMap
'    Set mappings = AdoTypeMappings.Default
'
'    Dim provider As IParameterProvider
'    Set provider = AdoParameterProvider.Create(mappings)
'
'    Dim baseCommand As IDbCommandBase
'    Set baseCommand = DbCommandBase.Create(provider)
'
'    Dim factory As IDbConnectionFactory
'    Set factory = New DbConnectionFactory 'the only other implementation is StubDbConnectionFactory, for unit tests.
'
'    Dim cmd As IDbCommand
'    Set cmd = AutoDbCommand.Create(connString, factory, baseCommand)
'
'    Dim results As ADODB.Recordset
'    Set results = cmd.Execute(SQLQuery)
'
'    Dim rows As Variant
'    rows = results.GetRows
'End Sub
'
'
'Private Sub DefaultDbCommandTest()
'    Dim ConnectionStringObject As SqliteConnectionString
'    Set ConnectionStringObject = SqliteConnectionString.Create(ThisWorkbook.Path, "SecureADODB.db")
'    Dim connString As String
'    connString = ConnectionStringObject.ADOConnectionString
'
'    Dim SQLQuery As String
'    SQLQuery = "SELECT * FROM categories WHERE category_id <= 3 AND section = 'machinery'"
'
'    Dim mappings As ITypeMap
'    Set mappings = AdoTypeMappings.Default
'
'    Dim provider As IParameterProvider
'    Set provider = AdoParameterProvider.Create(mappings)
'
'    Dim baseCommand As IDbCommandBase
'    Set baseCommand = DbCommandBase.Create(provider)
'
'    With DbConnection.Create(connString)
'        Dim cmd As IDbCommand
'        Set cmd = DefaultDbCommand.Create(.Self, baseCommand)
'
'        Dim results As ADODB.Recordset
'        Set results = cmd.Execute(SQLQuery)
'    End With
'
'    Dim rows As Variant
'    rows = results.GetRows
'End Sub
