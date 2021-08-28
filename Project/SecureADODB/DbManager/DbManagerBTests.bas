Attribute VB_Name = "DbManagerBTests"
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
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Factory Guard")
Private Sub ztcCreate_ThrowsGivenNullConnection()
    On Error Resume Next
    Dim sut As IDbManager: Set sut = DbManager.Create(Nothing, New StubDbCommandFactory)
    AssertExpectedError Assert, ErrNo.ObjectNotSetErr
End Sub


'@TestMethod("Factory Guard")
Private Sub ztcCreate_ThrowsGivenNullCommandFactory()
    On Error Resume Next
    Dim sut As IDbManager: Set sut = DbManager.Create(New StubDbConnection, Nothing)
    AssertExpectedError Assert, ErrNo.ObjectNotSetErr
End Sub


'@TestMethod("Create")
Private Sub ztcCommand_CreatesDbCommandWithFactory()
    Dim stubCommandFactory As StubDbCommandFactory
    Set stubCommandFactory = New StubDbCommandFactory
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(New StubDbConnection, stubCommandFactory)
    
    Dim Result As IDbCommand
    Set Result = sut.Command
    
    Assert.AreEqual 1, stubCommandFactory.CreateCommandInvokes
End Sub


'@TestMethod("Transaction")
Private Sub ztcCreate_StartsTransaction()
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    sut.Begin
    
    Assert.IsTrue stubConnection.DidBeginTransaction
End Sub


'@TestMethod("Transaction")
Private Sub ztcCommit_CommitsTransaction()
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Begin
    sut.Commit
    
    Assert.IsTrue stubConnection.DidCommitTransaction
End Sub


'@TestMethod("Transaction")
Private Sub ztcCommit_ThrowsIfAlreadyCommitted()
    On Error Resume Next
    
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Commit
    sut.Commit
    AssertExpectedError Assert, ErrNo.AdoInvalidTransactionErr
End Sub


'@TestMethod("Transaction")
Private Sub ztcCommit_ThrowsIfAlreadyRolledBack()
    On Error Resume Next
    
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Rollback
    sut.Commit
    AssertExpectedError Assert, ErrNo.AdoInvalidTransactionErr
End Sub


'@TestMethod("Transaction")
Private Sub ztcRollback_ThrowsIfAlreadyCommitted()
    On Error Resume Next
    
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Commit
    sut.Rollback
    AssertExpectedError Assert, ErrNo.AdoInvalidTransactionErr
End Sub


'@TestMethod("Transaction")
Private Sub ztcRollback_RollbacksTransaction()
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Begin
    sut.Rollback
    
    Assert.IsTrue stubConnection.DidRollBackTransaction
End Sub


'@TestMethod("Connection String")
Private Sub ztcBuildConnectionString_ThrowsGivenNullDatabaseType()
    On Error Resume Next
    Dim connString As String: connString = DbManager.BuildConnectionString(vbNullString)
    AssertExpectedError Assert, ErrNo.AdoConnectionStringErr
End Sub


'@TestMethod("Connection String")
Private Sub ztcBuildConnectionString_ThrowsGivenUnsupportedType()
    On Error Resume Next
    Dim connString As String: connString = DbManager.BuildConnectionString("Access")
    AssertExpectedError Assert, ErrNo.AdoConnectionStringErr
End Sub


'@TestMethod("Connection String")
Private Sub ztcBuildConnectionString_ValidatesDeafultCSVConnectionString()
    On Error GoTo TestFail
    
    Dim connString As String
    #If Win64 Then
        connString = "Driver=Microsoft Access Text Driver (*.txt, *.csv);DefaultDir=" + ThisWorkbook.Path + ";"
    #Else
        connString = "Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=" + ThisWorkbook.Path + ";"
    #End If
    Assert.AreEqual connString, DbManager.BuildConnectionString("csv"), "Default CSV connection string mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Connection String")
Private Sub ztcBuildConnectionString_ValidatesCSVConnectionString()
    On Error GoTo TestFail
    
    Dim connString As String
    #If Win64 Then
        connString = "Driver=Microsoft Access Text Driver (*.txt, *.csv);DefaultDir=C:\TMP;;"
    #Else
        connString = "Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=C:\TMP;;"
    #End If
    Assert.AreEqual connString, DbManager.BuildConnectionString("csv", "C:\TMP", "db.csv", ";"), "CSV connection string mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Connection String")
Private Sub ztcBuildConnectionString_ValidatesDeafultSQLiteConnectionString()
    On Error GoTo TestFail
    
    Dim connString As String
    connString = "Driver=SQLite3 ODBC Driver;Database=" + ThisWorkbook.Path + Application.PathSeparator + "SecureADODB.db;" + _
                 "SyncPragma=NORMAL;FKSupport=True;"
    Assert.AreEqual connString, DbManager.BuildConnectionString("sqlite"), "Default SQLite connection string mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Connection String")
Private Sub ztcBuildConnectionString_ValidatesSQLiteConnectionString()
    On Error GoTo TestFail
    
    Dim connString As String: connString = "Driver=SQLite3 ODBC Driver;Database=C:\TMP\db.db;_"
    Assert.AreEqual connString, DbManager.BuildConnectionString("sqlite", "C:\TMP", "db.db", "_"), "SQLite connection string mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Connection String")
Private Sub ztcBuildConnectionString_ValidatesRawConnectionString()
    On Error GoTo TestFail
    
    Dim connString As String: connString = "Driver=SQLite3 ODBC Driver;Database=C:\TMP\db.db;_"
    Assert.AreEqual connString, DbManager.BuildConnectionString(connString), "SQLite connection string mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub

