Attribute VB_Name = "DbManagerSampleCode"
'@Folder("-- DraftsTemplatesSnippets --")
'@IgnoreModule AssignmentNotUsed, EmptyModule, VariableNotUsed, ProcedureNotUsed, FunctionReturnValueDiscarded, FunctionReturnValueAlwaysDiscarded
Option Explicit


    
Private Sub DbManagerCSVTest()
    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim fileName As String: fileName = fso.GetBaseName(ThisWorkbook.Name) & ".csv"

    Dim tableName As String: tableName = fileName
    Dim SQLQuery As String: SQLQuery = "SELECT * FROM " & tableName & " WHERE id <= ? AND last_name <> 'machinery'"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.FromConnectionParameters("csv", ThisWorkbook.Path, fileName, vbNullString, True, LoggerTypeEnum.logPrivate)

    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Scalar:=False, Disconnected:=True, CacheSize:=10)
    
    rst.OpenRecordset SQLQuery, 45
End Sub


Private Sub DbManagerScalarCSVTest()
    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim fileName As String: fileName = fso.GetBaseName(ThisWorkbook.Name) & ".csv"

    Dim tableName As String: tableName = fileName
    Dim SQLQuery As String: SQLQuery = "SELECT last_name FROM " & tableName & " WHERE id = ? AND last_name <> 'machinery'"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.FromConnectionParameters("csv", ThisWorkbook.Path, fileName, vbNullString, True, LoggerTypeEnum.logPrivate)

    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Scalar:=True, Disconnected:=True, CacheSize:=10)
    
    Dim result As Variant
    result = rst.OpenScalar(SQLQuery, 45)
End Sub


Private Sub DbManagerSQLiteTest()
    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim fileName As String: fileName = fso.GetBaseName(ThisWorkbook.Name) & ".db"

    Dim tableName As String: tableName = "people"
    Dim SQLQuery As String: SQLQuery = "SELECT * FROM " & tableName & " WHERE id <= ? AND last_name <> 'machinery'"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.FromConnectionParameters("sqlite", ThisWorkbook.Path, fileName, vbNullString, True, LoggerTypeEnum.logPrivate)

    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Scalar:=False, Disconnected:=True, CacheSize:=10)
    
    rst.OpenRecordset SQLQuery, 45
End Sub


Private Sub DbManagerExTest()
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    Dim fileName As String
    fileName = fso.GetBaseName(ThisWorkbook.Name) & ".csv"
    Dim connString As String
    connString = DbManager.BuildConnectionString("csv", ThisWorkbook.Path, fileName)

    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & fileName & " WHERE id <= ? AND last_name <> 'machinery'"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.FromConnectionParameters("csv", ThisWorkbook.Path, fileName, vbNullString, True, LoggerTypeEnum.logPrivate)

    Dim Log As ILogger
    Set Log = dbm.LogController

    Dim conn As IDbConnection
    Set conn = dbm.Connection
    Dim connAdo As ADODB.Connection
    Set connAdo = conn.AdoConnection
    
    Dim cmd As IDbCommand
    Set cmd = dbm.Command
    Dim cmdAdo As ADODB.Command
    Set cmdAdo = cmd.AdoCommand(SQLQuery, 45)
    
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Scalar:=False, Disconnected:=True, CacheSize:=10)
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = rst.AdoRecordset
    
    rst.OpenRecordset SQLQuery, 45
End Sub


