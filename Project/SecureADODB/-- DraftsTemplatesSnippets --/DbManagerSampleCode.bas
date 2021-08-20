Attribute VB_Name = "DbManagerSampleCode"
'@Folder "SecureADODB.-- DraftsTemplatesSnippets --"
'@IgnoreModule AssignmentNotUsed, EmptyModule, VariableNotUsed, ProcedureNotUsed, FunctionReturnValueDiscarded, FunctionReturnValueAlwaysDiscarded
Option Explicit


Private Sub DbManagerCSVTest()
    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim FileName As String: FileName = fso.GetBaseName(ThisWorkbook.Name) & ".csv"

    Dim tableName As String: tableName = FileName
    Dim SQLQuery As String: SQLQuery = "SELECT * FROM " & tableName & " WHERE age >= ? AND country = 'South Korea'"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.FromConnectionParameters("csv", ThisWorkbook.Path, FileName, vbNullString, False, LoggerTypeEnum.logPrivate)

    '@Ignore IndexedDefaultMemberAccess
    Debug.Print dbm.Connection.AdoConnection.Properties("Transaction DDL").value
    
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Scalar:=False, Disconnected:=True, CacheSize:=10)
    
    Dim Result As ADODB.Recordset
    Set Result = rst.OpenRecordset(SQLQuery, 45)
End Sub


Private Sub DbManagerInvalidTypeTest()
    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim FileName As String: FileName = fso.GetBaseName(ThisWorkbook.Name) & ".csv"

    Dim tableName As String: tableName = FileName
    Dim SQLQuery As String: SQLQuery = "SELECT * FROM " & tableName & " WHERE age >= ? AND country = 'South Korea'"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.FromConnectionParameters("Driver=", ThisWorkbook.Path, FileName, vbNullString, True, LoggerTypeEnum.logPrivate)

    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Scalar:=False, Disconnected:=True, CacheSize:=10)
    
    Dim Result As ADODB.Recordset
    Set Result = rst.OpenRecordset(SQLQuery, 45)
End Sub


Private Sub DbManagerScalarCSVTest()
    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim FileName As String: FileName = fso.GetBaseName(ThisWorkbook.Name) & ".csv"

    Dim tableName As String: tableName = FileName
    Dim SQLQuery As String: SQLQuery = "SELECT * FROM " & tableName & " WHERE age >= ? AND country = 'South Korea' ORDER BY id DESC"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.FromConnectionParameters("csv", ThisWorkbook.Path, FileName, vbNullString, True, LoggerTypeEnum.logPrivate)

    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Scalar:=True, Disconnected:=True, CacheSize:=10)
    
    Dim Result As Variant
    Result = rst.OpenScalar(SQLQuery, 45)
End Sub


Private Sub DbManagerSQLiteTest()
    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim FileName As String: FileName = fso.GetBaseName(ThisWorkbook.Name) & ".db"

    Dim tableName As String: tableName = "people"
    Dim SQLQuery As String: SQLQuery = "SELECT * FROM " & tableName & " WHERE age >= ? AND country = 'South Korea'"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.FromConnectionParameters("sqlite", ThisWorkbook.Path, FileName, vbNullString, True, LoggerTypeEnum.logPrivate)

    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Scalar:=False, Disconnected:=True, CacheSize:=10)
    
    '@Ignore ImplicitDefaultMemberAccess, IndexedDefaultMemberAccess
    Debug.Print dbm.Connection.AdoConnection.Properties("Transaction DDL")
    
    Dim Result As ADODB.Recordset
    Set Result = rst.OpenRecordset(SQLQuery, 45)
End Sub


Private Sub DbManagerSQLiteInsertTest()
    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim FileName As String: FileName = fso.GetBaseName(ThisWorkbook.Name) & ".db"

    Dim tableName As String: tableName = "people_insert"
    Dim SQLQuery As String
    SQLQuery = "INSERT INTO " & tableName & " (id, first_name, last_name, age, gender, email, country, domain)" & _
               "VALUES (" & GenerateSerialID & ", 'first_name', 'last_name', 32, 'male', 'first_name.last_name@domain.com', 'Country', 'domain.com')"
               
    Dim dbm As IDbManager
    Set dbm = DbManager.FromConnectionParameters("sqlite", ThisWorkbook.Path, FileName, vbNullString, True, LoggerTypeEnum.logPrivate)
    
    Dim cmd As IDbCommand
    Set cmd = dbm.Command
    cmd.ExecuteNonQuery SQLQuery
    
    Dim conn As IDbConnection
    Set conn = dbm.Connection
    Dim RecordsAffected As Long
    RecordsAffected = conn.RecordsAffected
    Dim ExecuteStatus As ADODB.EventStatusEnum
    ExecuteStatus = conn.ExecuteStatus
End Sub


Private Sub DbManagerExTest()
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    Dim FileName As String
    FileName = fso.GetBaseName(ThisWorkbook.Name) & ".csv"
    Dim connString As String
    connString = DbManager.BuildConnectionString("csv", ThisWorkbook.Path, FileName)

    Dim tableName As String: tableName = FileName
    Dim SQLQuery As String: SQLQuery = "SELECT * FROM " & tableName & " WHERE age >= ? AND country = ?"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.FromConnectionParameters("csv", ThisWorkbook.Path, FileName, vbNullString, True, LoggerTypeEnum.logPrivate)

    Dim Log As ILogger
    Set Log = dbm.LogController

    Dim conn As IDbConnection
    Set conn = dbm.Connection
    Dim connAdo As ADODB.Connection
    Set connAdo = conn.AdoConnection
    
    Dim cmd As IDbCommand
    Set cmd = dbm.Command
    Dim cmdAdo As ADODB.Command
    Set cmdAdo = cmd.AdoCommand(SQLQuery, 45, "South Korea")
    
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Scalar:=False, Disconnected:=True, CacheSize:=10)
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = rst.AdoRecordset(SQLQuery, 45, "South Korea")
    
    Dim Result As ADODB.Recordset
    Set Result = rst.OpenRecordset(SQLQuery, 45, "South Korea")
End Sub


