VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A thin wrapper around an ADODB recordset."
'@Folder "SecureADODB.DbRecordset"
'@ModuleDescription("A thin wrapper around an ADODB recordset.")
'@PredeclaredId
Option Explicit

Implements IDbRecordset

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents AdoRecordset As ADODB.Recordset
Attribute AdoRecordset.VB_VarHelpID = -1
Private Type TRecordset
    
    OpenOptions As Long
End Type
Private this As TRecordset


'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As DbRecordset
Attribute Class.VB_Description = "Returns class reference"
    Set Class = DbRecordset
End Property


Public Function Create(ByVal cmd As IDbCommand, _
              Optional ByVal AsyncFetch As Boolean = False, _
              Optional ByVal CacheSize As Long = 10, _
              Optional ByVal CursorType As ADODB.CursorTypeEnum = ADODB.CursorTypeEnum.adOpenStatic, _
              Optional ByVal CursorLocation As ADODB.CursorLocationEnum = ADODB.CursorLocationEnum.adUseServer, _
              Optional ByVal AsynchOption As ADODB.ExecuteOptionEnum = 0) As IDbRecordset
              
    Errors.GuardNonDefaultInstance Me, DbRecordset

    Dim result As DbRecordset
    Set result = New DbRecordset
    result.CreInit cmd, AsyncFetch, CacheSize, CursorType, CursorLocation, AsynchOption
    Set Create = result
End Function


'@Description("Default constructor")
Friend Sub CreInit(ByVal cmd As IDbCommand, _
          Optional ByVal AsyncFetch As Boolean = False, _
          Optional ByVal CacheSize As Long = 10, _
          Optional ByVal CursorType As ADODB.CursorTypeEnum = ADODB.CursorTypeEnum.adOpenStatic, _
          Optional ByVal CursorLocation As ADODB.CursorLocationEnum = ADODB.CursorLocationEnum.adUseServer, _
          Optional ByVal AsynchOption As ADODB.ExecuteOptionEnum = 0)
Attribute CreInit.VB_Description = "Default constructor"
    
    Errors.GuardNullReference cmd, "SecureADODB.DbRecordset", "IDbCommand object is not set."
    
    Set AdoRecordset.source = cmd.AdoCommand
    AdoRecordset.CacheSize = CacheSize
    AdoRecordset.CursorLocation = CursorLocation
    AdoRecordset.CursorType = CursorType
    this.OpenOptions = AsynchOption Or (ADODB.ExecuteOptionEnum.adAsyncFetch And AsyncFetch)
End Sub


Private Sub Class_Initialize()
    Set AdoRecordset = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
    Set AdoRecordset = Nothing
End Sub


' https://docs.microsoft.com/en-us/sql/ado/guide/data/ado-event-instantiation-visual-basic
' https://docs.microsoft.com/en-us/sql/ado/guide/data/ado-event-handler-summary
' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/fetchprogress-event-ado
'
Private Sub AdoRecordset_FetchProgress(ByVal Progress As Long, ByVal MaxProgress As Long, ByRef adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Debug.Print "ADODB.Recordset fetch in progress. " & CStr(Progress) & " records of " & CStr(MaxProgress) & "have been retrived..."
End Sub


' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/fetchcomplete-event-ado
'
Private Sub AdoRecordset_FetchComplete(ByVal pError As ADODB.Error, ByRef adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Debug.Print "ADODB.Recordset fetch completed. "
    If adStatus = ADODB.EventStatusEnum.adStatusErrorsOccurred And pError Is Nothing Then
        Debug.Print "Error: " & pError.number, pError.description, pError.source, pError.source
    End If
End Sub


Private Property Get IDbRecordset_Self() As IDbRecordset
    Set IDbRecordset_Self = Me
End Property

Private Property Get IDbRecordset_AdoRecordset() As ADODB.Recordset
    Set IDbRecordset_AdoRecordset = AdoRecordset
End Property
