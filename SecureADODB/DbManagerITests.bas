Attribute VB_Name = "DbManagerITests"
'@Folder "SecureADODB.DbManager"
'@TestModule
'@IgnoreModule
Option Explicit
Option Private Module

#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If


'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub


'===================================================='
'===================== FIXTURES ====================='
'===================================================='


Private Function zfxGetConnectionString(ByVal TypeOrConnString As String) As String
    Dim fileExt As String: fileExt = IIf(TypeOrConnString = "csv", "csv", "db")
    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim fileName As String: fileName = fso.GetBaseName(ThisWorkbook.Name) & "." & fileExt
    
    zfxGetConnectionString = DbManager.BuildConnectionString(TypeOrConnString, ThisWorkbook.Path, fileName, vbNullString)
End Function


Private Function zfxGetDbManagerFromConnectionParameters(ByVal TypeOrConnString As String) As IDbManager
    Dim fileExt As String: fileExt = IIf(TypeOrConnString = "csv", "csv", "db")
    Dim fso As Scripting.FileSystemObject: Set fso = New Scripting.FileSystemObject
    Dim fileName As String: fileName = fso.GetBaseName(ThisWorkbook.Name) & "." & fileExt

    Dim dbm As IDbManager
    Set dbm = DbManager.FromConnectionParameters(TypeOrConnString, ThisWorkbook.Path, fileName, vbNullString, True, LoggerTypeEnum.logPrivate)
    Set zfxGetDbManagerFromConnectionParameters = dbm
End Function


Private Function zfxGetDbManagerFromConnectionString(ByVal TypeOrConnString As String) As IDbManager
    Dim dbm As IDbManager
    Set dbm = DbManager.FromConnectionParameters(TypeOrConnString) ' Use transactions and global Logger by default
    Set zfxGetDbManagerFromConnectionString = dbm
End Function


Private Function zfxGetSQLQuery(tableName As String) As String
    zfxGetSQLQuery = "SELECT * FROM " & tableName & " WHERE age >= 45 AND country = 'South Korea'"
End Function


Private Function zfxGetSQLQueryWithSingleParameter(tableName As String) As String
    zfxGetSQLQueryWithSingleParameter = "SELECT * FROM " & tableName & " WHERE age >= ? AND country = 'South Korea'"
End Function


Private Function zfxGetSQLQueryWithTwoParameters(tableName As String) As String
    zfxGetSQLQueryWithTwoParameters = "SELECT * FROM " & tableName & " WHERE age >= ? AND country = ?"
End Function


Private Function zfxGetCSVTableName() As String
    zfxGetCSVTableName = "SecureADODB.csv"
End Function


Private Function zfxGetSQLiteTableName() As String
    zfxGetSQLiteTableName = "people"
End Function


Private Function zfxGetParameterOne() As Variant
    zfxGetParameterOne = 45
End Function


Private Function zfxGetParameterTwo() As Variant
    zfxGetParameterTwo = "South Korea"
End Function


'===================================================='
'================= TESTING FIXTURES ================='
'===================================================='


'@TestMethod("Connection String")
Private Sub zfxGetConnectionString_VerifiesDefaultMockConnectionStrings()
    On Error GoTo TestFail
    
Arrange:
    Dim CSVString As String
    CSVString = "Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=" + ThisWorkbook.Path + ";"
    Dim SQLiteString As String
    SQLiteString = "Driver={SQLite3 ODBC Driver};Database=" + ThisWorkbook.Path + Application.PathSeparator + "SecureADODB.db;" + _
                   "SyncPragma=NORMAL;LongNames=True;NoCreat=True;FKSupport=True;OEMCP=True;"
Act:

Assert:
    Assert.AreEqual CSVString, DbManager.BuildConnectionString("csv"), "Default CSV string mismatch"
    Assert.AreEqual SQLiteString, DbManager.BuildConnectionString("sqlite")

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("Connection String")
Private Sub zfxGetDbManagerFromConnectionParameters_ThrowsGivenInvalidConnectionString()
    On Error Resume Next
    Dim TypeOrConnString As String: TypeOrConnString = "Driver={SQLite3 ODBC Driver};Database=C:\TMP\db.db;"
    Dim dbm As IDbManager: Set dbm = zfxGetDbManagerFromConnectionParameters(TypeOrConnString)
    AssertExpectedError Assert, ErrNo.AdoConnectionStringError
End Sub


'===================================================='
'================ TEST MOCK DATABASE ================'
'===================================================='


'@TestMethod("DbManager.Command")
Private Sub ztiDbManagerCommand_VerifiesAdoCommand()
    On Error GoTo TestFail
    
Arrange:
    Dim dbm As IDbManager: Set dbm = DbManager.FromConnectionParameters(zfxGetConnectionString("sqlite"))
    Dim SQLQuery2P As String: SQLQuery2P = zfxGetSQLQueryWithTwoParameters(zfxGetSQLiteTableName)
Act:
    Dim cmdAdo As ADODB.Command
    Set cmdAdo = dbm.Command.AdoCommand(SQLQuery2P, zfxGetParameterOne, zfxGetParameterTwo)
Assert:
    Assert.IsNotNothing cmdAdo.ActiveConnection, "ActiveConnection of the Command object is not set."
    Assert.AreEqual ADODB.ObjectStateEnum.adStateOpen, cmdAdo.ActiveConnection.State, "ActiveConnection of the Command object is not open."
    Assert.IsTrue cmdAdo.Prepared, "Prepared property of the Command object not set."
    Assert.AreEqual 2, cmdAdo.Parameters.Count, "Command should have two parameters set."
    Assert.AreEqual ADODB.DataTypeEnum.adInteger, cmdAdo.Parameters.Item(0).Type, "Param #1 type should be adInteger."
    Assert.AreEqual 45, cmdAdo.Parameters.Item(0).value, "Param #1 value should be 45."
    Assert.AreEqual ADODB.DataTypeEnum.adVarWChar, cmdAdo.Parameters.Item(1).Type, "Param #2 type should be adVarWChar."
    Assert.AreEqual "South Korea", cmdAdo.Parameters.Item(1).value, "Param #2 value should be South Korea."
    Assert.AreNotEqual vbNullString, cmdAdo.CommandText
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("DbManager.Recordset")
Private Sub ztiDbManagerCommand_VerifiesAdoRecordsetDefaultDisconnectedArray()
    On Error GoTo TestFail
    
Arrange:
    Dim dbm As IDbManager: Set dbm = DbManager.FromConnectionParameters(zfxGetConnectionString("sqlite"))
    Dim SQLQuery2P As String: SQLQuery2P = zfxGetSQLQueryWithTwoParameters(zfxGetSQLiteTableName)
Act:
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = dbm.Recordset.AdoRecordset(SQLQuery2P, zfxGetParameterOne, zfxGetParameterTwo)
Assert:
    Assert.IsNotNothing rstAdo.ActiveConnection, "ActiveConnection of the Recordset object is not set."
    Assert.IsNotNothing rstAdo.ActiveCommand, "ActiveCommand of the Recordset object is not set."
    Assert.IsFalse IsFalsy(rstAdo.source), "The Source property of the Recordset object is not set."
    Assert.AreEqual ADODB.CursorTypeEnum.adOpenStatic, rstAdo.CursorType, "The CursorType of the Recordset object should be adOpenStatic."
    Assert.AreEqual ADODB.CursorLocationEnum.adUseClient, rstAdo.CursorLocation, "The CursorLocation of the Recordset object should be adUseClient."
    Assert.AreNotEqual 1, rstAdo.MaxRecords, "The MaxRecords of the Recordset object should not be set to 1 for a regular Recordset."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("DbManager.Recordset")
Private Sub ztiDbManagerCommand_VerifiesAdoRecordsetDisconnectedScalar()
    On Error GoTo TestFail
    
Arrange:
    Dim dbm As IDbManager: Set dbm = DbManager.FromConnectionParameters(zfxGetConnectionString("sqlite"))
    Dim SQLQuery2P As String: SQLQuery2P = zfxGetSQLQueryWithTwoParameters(zfxGetSQLiteTableName)
Act:
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Scalar:=True, CacheSize:=15)
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = rst.AdoRecordset(SQLQuery2P, zfxGetParameterOne, zfxGetParameterTwo)
Assert:
    Assert.AreEqual 1, rstAdo.MaxRecords, "The MaxRecords of the Recordset object should be set to 1 for a scalar query."
    Assert.AreEqual 15, rstAdo.CacheSize, "The CacheSize of the Recordset object should be set to 15."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("DbManager.Recordset")
Private Sub ztiDbManagerCommand_VerifiesAdoRecordsetOnlineArray()
    On Error GoTo TestFail
    
Arrange:
    Dim dbm As IDbManager: Set dbm = DbManager.FromConnectionParameters(zfxGetConnectionString("sqlite"), , , , False)
    Dim SQLQuery2P As String: SQLQuery2P = zfxGetSQLQueryWithTwoParameters(zfxGetSQLiteTableName)
Act:
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Disconnected:=False)
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = rst.AdoRecordset(SQLQuery2P, zfxGetParameterOne, zfxGetParameterTwo)
Assert:
    Assert.AreEqual ADODB.CursorTypeEnum.adOpenForwardOnly, rstAdo.CursorType, "The CursorType of the Recordset object should be adOpenForwardOnly."
    Assert.AreEqual ADODB.CursorLocationEnum.adUseServer, rstAdo.CursorLocation, "The CursorLocation of the Recordset object should be adUseServer."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub

