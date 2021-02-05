VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelConnectionString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "-- DraftsTemplatesSnippets --.DbConnectionString"
'@PredeclaredId
'@IgnoreModule

Option Explicit


Private Type TExcelConnectionString
    Driver As String
    Options As String
    Path As String
    Name As String
    pathName As String
    provider As String
End Type
Private this As TExcelConnectionString


'@Description("Returns instance reference")
'@Ignore ProcedureNotUsed
Public Property Get Self() As ExcelConnectionString
Attribute Self.VB_Description = "Returns instance reference"
    Set Self = Me
End Property


'@Description("Returns class reference")
'@Ignore ProcedureNotUsed
Public Property Get Class() As ExcelConnectionString
Attribute Class.VB_Description = "Returns class reference"
    Set Class = ExcelConnectionString
End Property


'@Description("Default factory")
Public Function Create(Optional ByVal dbPath As String, _
                       Optional ByVal dbName As String, _
                       Optional ByVal Driver As String = vbNullString, _
                       Optional ByVal Options As String = vbNullString, _
                       Optional ByVal provider As String = vbNullString) As ExcelConnectionString
Attribute Create.VB_Description = "Default factory"
    Dim result As ExcelConnectionString
    Set result = New ExcelConnectionString
    result.Init dbPath, dbName, Driver, Options, provider
    Set Create = result
End Function


'@Description("Custom constructor")
Public Sub Init(ByVal dbPath As String, ByVal dbName As String, ByVal Driver As String, ByVal Options As String, ByVal provider As String)
Attribute Init.VB_Description = "Custom constructor"
    Dim dbPathName As String
    dbPathName = dbPath & Application.PathSeparator & dbName
    Guard.FileNotExist dbPathName
    
    this.Name = dbName
    this.Path = dbPath
    this.pathName = dbPathName
    
    If Driver = vbNullString Then
        this.Driver = "{SQLite3 ODBC Driver}"
    Else
        this.Driver = Driver
    End If
    If Options = vbNullString Then
        this.Options = "SyncPragma=NORMAL;LongNames=True;NoCreat=True;FKSupport=True;OEMCP=True;"
    Else
        this.Options = Options
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
    ADOConnectionString = Join(Array("Driver=", this.Driver, ";", "Database=", this.pathName, ";", this.Options), vbNullString)
End Property


''@Description("Returns connection string for an Excel QueryTable object")
''@Ignore ProcedureNotUsed
'Public Property Get QTConnectionString() As String
'    QTConnectionString = this.provider & ";" & Me.ADOConnectionString
'End Property

'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=F:\Archive\Business\FID\PolMaFID\Drafts\Knowledge Management System\VBA\Samples\RDVBA-examples\SecureADODB.xls;Persist Security Info=False;Extended Properties="Excel 12.0;HDR=Yes"
