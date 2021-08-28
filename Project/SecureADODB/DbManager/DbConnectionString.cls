VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbConnectionString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Helper routines for building connection strings"
'@Folder "SecureADODB.DbManager"
'@ModuleDescription "Helper routines for building connection strings"
'@PredeclaredId
''''
'''' The module incorporates routines, which facilitate construction of
'''' connection strings for the ADODB library. Presently, a distinction is made
'''' between file-based databases, such as sqlite, csv, xls, etc., and network
'''' based databases. A file-based database is accessed based on its type and
'''' file pathname.
'''' CreateFileDB/InitFileDB pair is used for file based databases.
''''
Option Explicit
Option Compare Text

Private Type TDbConnectionString
    DbType As String
    DbPath As String
    Options As String
    Driver As String
    ConnectionString As String
End Type
Private this As TDbConnectionString


'''' Factory for file-based databases.
''''
'''' Args:
''''   FileType (string, optional, "sqlite"):
''''     Type of the database: "sqlite", "csv", "xls"
''''
''''   FileName (string, optional, ""):
''''     Database file name. If not provided, Thisworkbook.Name (without
''''     extension) will be used. Extension is added based on the database type:
''''       "sqlite" - "db" or "sqlite"
''''       "csv"    - "xsv" or "csv"
''''       "xls"    - "xls
''''
''''   Driver (variant, optional, Empty):
''''     Database driver. If Empty, default values is selected based on
''''     the database type (see constructor code for details).
''''
''''   Options (variant, optional, Empty):
''''     Connection options. See above for details.
''''
''''  N.B.: "xls" backend support is not currently implemented
''''
'''' Examples:
''''   >>> ?DbConnectionString.CreateFileDB("sqlite").ConnectionString
''''   "Driver=SQLite3 ODBC Driver;Database=<Thisworkbook.Path>\SecureADODB.db;SyncPragma=NORMAL;FKSupport=True;"
''''
''''   >>> ?DbConnectionString.CreateFileDB("sqlite").QTConnectionString
''''   "OLEDB;Driver=SQLite3 ODBC Driver;Database=<Thisworkbook.Path>\SecureADODB.db;SyncPragma=NORMAL;FKSupport=True;"
''''
''''   >>> ?DbConnectionString.CreateFileDB("csv").ConnectionString
''''   "Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=<Thisworkbook.Path>\SecureADODB.csv;"
''''
''''   >>> ?DbConnectionString.CreateFileDB("xls").ConnectionString
''''   NotImplementedErr
''''
'@Description "Factory for file-based databases"
Public Function CreateFileDB(Optional ByVal FileType As String = "sqlite", _
                             Optional ByVal FileName As String = vbNullString, _
                             Optional ByVal Driver As Variant = Empty, _
                             Optional ByVal Options As Variant = Empty _
                            ) As DbConnectionString
Attribute CreateFileDB.VB_Description = "Factory for file-based databases"
    Dim Instance As DbConnectionString
    Set Instance = New DbConnectionString
    Instance.InitFileDB FileType, FileName, Driver, Options
    Set CreateFileDB = Instance
End Function


'@Description "Constructor for file-based databases"
Friend Sub InitFileDB(Optional ByVal FileType As String = "sqlite", _
                      Optional ByVal FileName As String = vbNullString, _
                      Optional ByVal Driver As String = Empty, _
                      Optional ByVal Options As Variant = Empty)
Attribute InitFileDB.VB_Description = "Constructor for file-based databases"
    With this
        .DbType = LCase$(FileType)
        .Driver = Driver
        .Options = Options
        Select Case .DbType
            Case "sqlite"
                If IsEmpty(Options) Then
                    .Driver = "SQLite3 ODBC Driver"
                End If
                If IsEmpty(Options) Then
                    .Options = "SyncPragma=NORMAL;FKSupport=True;"
                End If
                .DbPath = VerifyOrGetDefaultPath(FileName, Array("db", "sqlite"))
                .ConnectionString = "Driver=" + .Driver + ";" + _
                                    "Database=" + .DbPath + ";" + _
                                    .Options
            Case "worksheet", "wsheet"
                .DbPath = VerifyOrGetDefaultPath(FileName, Array("xls"))
                .ConnectionString = .DbPath
            Case "csv"
                .DbPath = VerifyOrGetDefaultPath(FileName, Array("xsv", "csv"))
                #If Win64 Then
                    .Driver = "Microsoft Access Text Driver (*.txt, *.csv)"
                #Else
                    .Driver = "{Microsoft Text Driver (*.txt; *.csv)}"
                #End If
                .ConnectionString = "Driver=" + .Driver + ";" + _
                                    "DefaultDir=" + .DbPath + ";"
            Case Else
                .ConnectionString = vbNullString
        End Select
    End With

    If this.ConnectionString = vbNullString Then
        Dim errorDetails As TError
        With errorDetails
            .Number = ErrNo.NotImplementedErr
            .Name = "NotImplementedErr"
            .Source = "DbConnectionString"
            .Description = "Unsupported backend: " & FileType
            .Message = .Description
        End With
        RaiseError errorDetails
    End If
End Sub


Public Property Get ConnectionString() As String
    ConnectionString = this.ConnectionString
End Property


Public Property Get QTConnectionString() As String
    QTConnectionString = "OLEDB;" & this.ConnectionString
End Property
