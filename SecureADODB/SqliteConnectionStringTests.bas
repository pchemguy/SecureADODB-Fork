Attribute VB_Name = "SqliteConnectionStringTests"
'@Folder("-- DraftsTemplatesSnippets --")
'@TestModule
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module


Private Assert As Object


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
    'this method runs once per module.
    Set Assert = Nothing
End Sub


'@TestInitialize
'@Ignore EmptyMethod
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub


'@TestCleanup
'@Ignore EmptyMethod
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'@TestMethod("SqliteConnectionString.Self")
Private Sub Self_CheckAvailability()
    On Error GoTo TestFail
    
Arrange:
    Dim instanceVar As Object: Set instanceVar = SqliteConnectionString.Create(ThisWorkbook.path, "SecureADODB.db")
Act:
    Dim selfVar As Object: Set selfVar = instanceVar.Self
Assert:
    Assert.AreEqual TypeName(instanceVar), TypeName(selfVar), "Error: type mismatch: " & TypeName(selfVar) & " type."
    Assert.AreSame instanceVar, selfVar, "Error: bad Self pointer"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("SqliteConnectionString.ClassName")
Private Sub ClassName_CheckAvailability()
    On Error GoTo TestFail
    
Arrange:
    Dim classVar As Object: Set classVar = SqliteConnectionString
Act:
    Dim classNameVar As Object: Set classNameVar = classVar.Create(ThisWorkbook.path, "SecureADODB.db").ClassName
Assert:
    Assert.AreEqual TypeName(classVar), TypeName(classNameVar), "Error: type mismatch: " & TypeName(classNameVar) & " type."
    Assert.AreSame classVar, classNameVar, "Error: bad Class pointer"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("SqliteConnectionString.Create")
Private Sub Create_ThrowsIfDbFileNotExist()
    On Error Resume Next
    Guard.FileNotExist vbNullString
    AssertExpectedError ErrNo.FileNotFoundErr
End Sub


'@TestMethod("SqliteConnectionString.Create")
Private Sub Create_Pass()
    On Error Resume Next
    Guard.FileNotExist ThisWorkbook.name
    AssertExpectedError ErrNo.PassedNoErr
End Sub


'@TestMethod("SqliteConnectionString.Create")
Private Sub Create_ValidatesConnectionString()
    On Error GoTo TestFail
    
Arrange:
    Dim ConnectionStringText As String
    ConnectionStringText = "Driver={SQLite3 ODBC Driver};Database=" _
                           & ThisWorkbook.path & Application.PathSeparator _
                           & "SecureADODB.db;SyncPragma=NORMAL;LongNames=True;NoCreat=True;FKSupport=True;OEMCP=True;"
    Dim ErrorMessage As String
    ErrorMessage = "Error: connection string mismatch - expected vs. actual -" & vbNewLine
Act:
    Dim ConnectionStringObject As SqliteConnectionString
    Set ConnectionStringObject = SqliteConnectionString.Create(ThisWorkbook.path, "SecureADODB.db")
Assert:
    Assert.AreEqual ConnectionStringText, ConnectionStringObject.ADOConnectionString, _
                    ErrorMessage & ConnectionStringText & vbNewLine & ConnectionStringObject.ADOConnectionString
    Assert.AreEqual "OLEDB;" & ConnectionStringText, ConnectionStringObject.QTConnectionString, _
                    ErrorMessage & "OLEDB;" & ConnectionStringText & vbNewLine & ConnectionStringObject.QTConnectionString

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


