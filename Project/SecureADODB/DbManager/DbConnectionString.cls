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


Public Function CreateFileDB(Optional ByVal FileType As String = "sqlite", _
                             Optional ByVal FilePath As String = vbNullString, _
                             Optional ByVal FileName As String = vbNullString, _
                             Optional ByVal Driver As Variant = Empty, _
                             Optional ByVal Options As Variant = Empty _
                             ) As DbConnectionString
    Dim Instance As DbConnectionString
    Set Instance = New DbConnectionString
    Instance.InitFileDB FileType, FilePath, FileName, Driver, Options
    Set CreateFileDB = Instance
End Function


Friend Sub InitFileDB(Optional ByVal FileType As String = "sqlite", _
                      Optional ByVal FilePath As String = vbNullString, _
                      Optional ByVal FileName As String = vbNullString, _
                      Optional ByVal Driver As String = Empty, _
                      Optional ByVal Options As Variant = Empty)
    With this
        .DbPath = FilePath & Application.PathSeparator & FileName
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
                .DbPath = VerifyOrGetDefaultPath(.DbPath, Array("db", "sqlite"))
                .ConnectionString = "Driver=" + .Driver + ";" + _
                                    "Database=" + .DbPath + ";" + _
                                    .Options
            Case "worksheet", "wsheet"
                .DbPath = VerifyOrGetDefaultPath(.DbPath, Array("xls"))
                .ConnectionString = .DbPath
            Case "csv"
                .DbPath = VerifyOrGetDefaultPath(.DbPath, Array("xsv", "csv"))
                #If Win64 Then
                    .Driver = "Microsoft Access Text Driver (*.txt, *.csv)"
                #Else
                    .Driver = "{Microsoft Text Driver (*.txt; *.csv)}"
                #End If
                .ConnectionString = "Driver=" + .Driver + ";" + _
                                    "DefaultDir=" + .DbPath + ";" + _
                                    .Options
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
