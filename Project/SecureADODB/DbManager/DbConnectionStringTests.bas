Attribute VB_Name = "DbConnectionStringTests"
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

'@TestMethod("ConnectionString")
Private Sub ztcConnectionString_ValidatesDefaultSQLiteString()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = "Driver=SQLite3 ODBC Driver;Database=" & _
                VerifyOrGetDefaultPath(vbNullString, Array("db", "sqlite")) & _
               ";SyncPragma=NORMAL;FKSupport=True;"
Act:
    Dim ActualADO As String
    ActualADO = DbConnectionString.CreateFileDB("sqlite").ConnectionString
    Dim ActualQT As String
    ActualQT = DbConnectionString.CreateFileDB("sqlite").QTConnectionString
Assert:
    Assert.AreEqual Expected, ActualADO, "Default SQLite ADO ConnectionString mismatch"
    Assert.AreEqual "OLEDB;" & Expected, ActualQT, "Default SQLite QT ConnectionString mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ConnectionString")
Private Sub ztcConnectionString_ValidatesDefaultCSVString()
    On Error GoTo TestFail

Arrange:
    Dim Driver As String
    #If Win64 Then
        Driver = "Microsoft Access Text Driver (*.txt, *.csv)"
    #Else
        Driver = "{Microsoft Text Driver (*.txt; *.csv)}"
    #End If
    Dim Expected As String
    Expected = "Driver=" + Driver + ";" + "DefaultDir=" + _
               VerifyOrGetDefaultPath(vbNullString, Array("xsv", "csv")) + ";"
Act:
    Dim ActualADO As String
    ActualADO = DbConnectionString.CreateFileDB("csv").ConnectionString
    Dim ActualQT As String
    ActualQT = DbConnectionString.CreateFileDB("csv").QTConnectionString
Assert:
    Assert.AreEqual Expected, ActualADO, "Default CSV ADO ConnectionString mismatch"
    Assert.AreEqual "OLEDB;" & Expected, ActualQT, "Default CSV QT ConnectionString mismatch"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ConnectionString")
Private Sub ztcConnectionString_ThrowsForXLSBackend()
    On Error Resume Next
    Dim ConnectionString As String
    ConnectionString = DbConnectionString.CreateFileDB("xls").ConnectionString
    AssertExpectedError Assert, ErrNo.NotImplementedErr
End Sub

