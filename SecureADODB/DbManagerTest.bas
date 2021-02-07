Attribute VB_Name = "DbManagerTest"
'@Folder "SecureADODB.DbManager"
'@TestModule
'@IgnoreModule
Option Explicit
Option Private Module

Private Const ExpectedError As Long = SecureADODBCustomError

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


'@TestMethod("Factory Guard")
Private Sub Create_ThrowsIfNotInvokedFromDefaultInstance()
    On Error Resume Next
    Dim sutObject As DbManager
    Set sutObject = New DbManager
    Dim sutInterface As IDbManager
    Set sutInterface = sutObject.Create(New StubDbConnection, New StubDbCommandFactory)
    AssertExpectedError Assert, ErrNo.NonDefaultInstanceErr
End Sub


'@TestMethod("Factory Guard")
Private Sub Create_ThrowsGivenNullConnection()
    On Error Resume Next
    Dim sut As IDbManager: Set sut = DbManager.Create(Nothing, New StubDbCommandFactory)
    AssertExpectedError Assert, ErrNo.ObjectNotSetErr
End Sub


'@TestMethod("Factory Guard")
Private Sub Create_ThrowsGivenNullCommandFactory()
    On Error Resume Next
    Dim sut As IDbManager: Set sut = DbManager.Create(New StubDbConnection, Nothing)
    AssertExpectedError Assert, ErrNo.ObjectNotSetErr
End Sub


'@TestMethod("Transaction")
Private Sub Command_CreatesDbCommandWithFactory()
    Dim stubCommandFactory As StubDbCommandFactory
    Set stubCommandFactory = New StubDbCommandFactory
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(New StubDbConnection, stubCommandFactory)
    
    Dim result As IDbCommand
    Set result = sut.Command
    
    Assert.AreEqual 1, stubCommandFactory.CreateCommandInvokes
End Sub


'@TestMethod("Transaction")
Private Sub Create_StartsTransaction()
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    
    Assert.IsTrue stubConnection.DidBeginTransaction
End Sub


'@TestMethod("Transaction")
Private Sub Commit_CommitsTransaction()
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Commit
    
    Assert.IsTrue stubConnection.DidCommitTransaction
End Sub


'@TestMethod("Transaction")
Private Sub Commit_ThrowsIfAlreadyCommitted()
    On Error GoTo TestFail
    
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Commit
    On Error GoTo CleanFail
    sut.Commit
    On Error GoTo 0

CleanFail:
    If Err.number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub


'@TestMethod("Transaction")
Private Sub Commit_ThrowsIfAlreadyRolledBack()
    On Error GoTo TestFail
    
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Rollback
    On Error GoTo CleanFail
    sut.Commit
    On Error GoTo 0

CleanFail:
    If Err.number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub


'@TestMethod("Transaction")
Private Sub Rollback_ThrowsIfAlreadyCommitted()
    On Error GoTo TestFail
    
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Commit
    On Error GoTo CleanFail
    sut.Rollback
    On Error GoTo 0

CleanFail:
    If Err.number = ExpectedError Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub


'@TestMethod("Connection String")
Private Sub BuildConnectionString_ThrowsGivenNullDatabaseType()
    On Error Resume Next
    Dim connString As String: connString = DbManager.BuildConnectionString(vbNullString)
    AssertExpectedError Assert, ErrNo.EmptyStringErr
End Sub


'@TestMethod("Connection String")
Private Sub BuildConnectionString_UnsupportedType()
    On Error GoTo TestFail
    Assert.AreEqual vbNullString, DbManager.BuildConnectionString("Access")

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("Connection String")
Private Sub BuildConnectionString_DeafultCSVConnectionString()
    On Error GoTo TestFail
    
    Dim connString As String
    connString = "Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=" + ThisWorkbook.Path + ";"
    Assert.AreEqual connString, DbManager.BuildConnectionString("csv")

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("Connection String")
Private Sub BuildConnectionString_CSVConnectionString()
    On Error GoTo TestFail
    
    Dim connString As String
    connString = "Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=C:\TMP;;"
    Assert.AreEqual connString, DbManager.BuildConnectionString("csv", "C:\TMP", "db.csv", ";")

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("Connection String")
Private Sub BuildConnectionString_DeafultSQLiteConnectionString()
    On Error GoTo TestFail
    
    Dim connString As String
    connString = "Driver={SQLite3 ODBC Driver};Database=" + ThisWorkbook.Path + Application.PathSeparator + "SecureADODB.db;" + _
                 "SyncPragma=NORMAL;LongNames=True;NoCreat=True;FKSupport=True;OEMCP=True;"
    Assert.AreEqual connString, DbManager.BuildConnectionString("sqlite")

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("Connection String")
Private Sub BuildConnectionString_SQLiteConnectionString()
    On Error GoTo TestFail
    
    Dim connString As String
    connString = "Driver={SQLite3 ODBC Driver};Database=C:\TMP\db.db;_"
    Assert.AreEqual connString, DbManager.BuildConnectionString("sqlite", "C:\TMP", "db.db", "_")

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("Connection String")
Private Sub BuildConnectionString_RawConnectionString()
    On Error GoTo TestFail
    
    Dim connString As String
    connString = "Driver={SQLite3 ODBC Driver};Database=C:\TMP\db.db;_"
    Assert.AreEqual connString, DbManager.BuildConnectionString(connString)

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


