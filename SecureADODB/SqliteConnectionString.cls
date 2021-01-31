VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqliteConnectionString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("-- DraftsTemplatesSnippets --")
'@PredeclaredId
Option Explicit


Private Type TSqliteConnectionString
    driver As String
    options As String
    path As String
    name As String
    pathName As String
    provider As String
End Type
Private this As TSqliteConnectionString


'@Description("Returns instance reference")
'@Ignore ProcedureNotUsed
Public Property Get Self() As SqliteConnectionString
Attribute Self.VB_Description = "Returns instance reference"
    Set Self = Me
End Property


'@Description("Returns class reference")
'@Ignore ProcedureNotUsed
Public Property Get Class() As SqliteConnectionString
    Set Class = SqliteConnectionString
End Property


'@Description("Default factory")
Public Function Create(ByVal dbPath As String, _
                       ByVal dbName As String, _
              Optional ByVal driver As String = vbNullString, _
              Optional ByVal options As String = vbNullString, _
              Optional ByVal provider As String = vbNullString) As SqliteConnectionString
Attribute Create.VB_Description = "Default factory"
    Dim result As SqliteConnectionString
    Set result = New SqliteConnectionString
    result.Init dbPath, dbName, driver, options, provider
    Set Create = result
End Function


'@Description("Custom constructor")
Public Sub Init(ByVal dbPath As String, ByVal dbName As String, ByVal driver As String, ByVal options As String, ByVal provider As String)
Attribute Init.VB_Description = "Custom constructor"
    Dim dbPathName As String
    dbPathName = dbPath & Application.PathSeparator & dbName
    Guard.FileNotExist dbPathName
    
    this.name = dbName
    this.path = dbPath
    this.pathName = dbPathName
    
    If driver = vbNullString Then
        this.driver = "{SQLite3 ODBC Driver}"
    Else
        this.driver = driver
    End If
    If options = vbNullString Then
        this.options = "SyncPragma=NORMAL;LongNames=True;NoCreat=True;FKSupport=True;OEMCP=True;"
    Else
        this.options = options
    End If
    If provider = vbNullString Then
        this.provider = "OLEDB"
    Else
        this.provider = provider
    End If
End Sub


'@Description("Returns connection string for an ADODB object")
Public Property Get ADOConnectionString() As String
Attribute ADOConnectionString.VB_Description = "Returns connection string for an ADODB object"
    ADOConnectionString = Join(Array("Driver=", this.driver, ";", "Database=", this.pathName, ";", this.options), vbNullString)
End Property


'@Description("Returns connection string for an Excel QueryTable object")
'@Ignore ProcedureNotUsed
Public Property Get QTConnectionString() As String
Attribute QTConnectionString.VB_Description = "Returns connection string for an Excel QueryTable object"
    QTConnectionString = this.provider & ";" & Me.ADOConnectionString
End Property
