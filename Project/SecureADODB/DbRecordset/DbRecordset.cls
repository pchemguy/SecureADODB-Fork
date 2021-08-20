VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A thin wrapper around an ADODB recordset."
'@Folder "SecureADODB.DbRecordset"
'@ModuleDescription("A thin wrapper around an ADODB recordset.")
'@PredeclaredId
'@Exposed
Option Explicit

Implements IDbRecordset

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents AdoRecordset As ADODB.Recordset
Attribute AdoRecordset.VB_VarHelpID = -1
Private Type TRecordset
    cmd As IDbCommand
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
              Optional ByVal Scalar As Boolean = False, _
              Optional ByVal Disconnected As Boolean = True, _
              Optional ByVal CacheSize As Long = 10, _
              Optional ByVal CursorType As ADODB.CursorTypeEnum = -1, _
              Optional ByVal AsyncMode As Boolean = False, _
              Optional ByVal AsyncOption As ADODB.ExecuteOptionEnum = 0) As IDbRecordset
              
    Guard.NonDefaultInstance Me

    Dim Result As DbRecordset
    Set Result = New DbRecordset
    Result.CreInit cmd, Scalar, Disconnected, CacheSize, CursorType, AsyncMode, AsyncOption
    Set Create = Result
End Function


'@Description("Default constructor")
Friend Sub CreInit(ByVal cmd As IDbCommand, _
          Optional ByVal Scalar As Boolean = False, _
          Optional ByVal Disconnected As Boolean = True, _
          Optional ByVal CacheSize As Long = 10, _
          Optional ByVal CursorType As ADODB.CursorTypeEnum = -1, _
          Optional ByVal AsyncMode As Boolean = False, _
          Optional ByVal AsyncOption As ADODB.ExecuteOptionEnum = 0)
Attribute CreInit.VB_Description = "Default constructor"
    
    Set this.cmd = cmd
    If Scalar Then AdoRecordset.MaxRecords = 1
    AdoRecordset.CursorLocation = IIf(Disconnected, ADODB.CursorLocationEnum.adUseClient, ADODB.CursorLocationEnum.adUseServer)
    AdoRecordset.CacheSize = CacheSize
    If CursorType > 0 Then
        AdoRecordset.CursorType = CursorType
    Else
        AdoRecordset.CursorType = IIf(Disconnected, ADODB.CursorTypeEnum.adOpenStatic, ADODB.CursorTypeEnum.adOpenForwardOnly)
    End If
    this.OpenOptions = AsyncOption Or (ADODB.ExecuteOptionEnum.adAsyncFetch And AsyncMode)
End Sub


Private Sub Class_Initialize()
    Set AdoRecordset = New ADODB.Recordset
End Sub


Private Sub Class_Terminate()
    Set AdoRecordset = Nothing
End Sub


Private Property Get IDbRecordset_Self() As IDbRecordset
    Set IDbRecordset_Self = Me
End Property


'@Description("Returns class reference")
Private Property Get IDbRecordset_Class() As DbRecordset
Attribute IDbRecordset_Class.VB_Description = "Returns class reference"
    Set IDbRecordset_Class = DbRecordset
End Property


Private Function IDbRecordset_AdoRecordset(ByVal sql As String, ParamArray args() As Variant) As ADODB.Recordset
    If sql <> vbNullString And AdoRecordset.ActiveCommand Is Nothing Then
        Dim localArgs() As Variant
        If IsArray(args(0)) Then
            localArgs = args(0)
        Else
            localArgs = args
        End If
        
        Dim AdoCommand As ADODB.Command
        Set AdoCommand = this.cmd.AdoCommand(sql, localArgs)
        Set AdoRecordset.Source = AdoCommand
    End If
    Set IDbRecordset_AdoRecordset = AdoRecordset
End Function


' Execute and ExecuteScalar can be combined into one method returning Variant, where distinction
' is made based on the AdoRecordset.MaxRecords value (1 - Scalar, regular otherwise)
Private Function IDbRecordset_OpenRecordset(ByVal sql As String, ParamArray args() As Variant) As ADODB.Recordset
    '@Ignore FunctionReturnValueDiscarded
    IDbRecordset_AdoRecordset sql, args
    
    On Error GoTo RecordsetOpenError
    AdoRecordset.Open Options:=this.OpenOptions
    On Error GoTo 0
    If AdoRecordset.CursorLocation = adUseClient Then Set AdoRecordset.ActiveConnection = Nothing
    Set IDbRecordset_OpenRecordset = AdoRecordset
    Exit Function
    
RecordsetOpenError:
    Err.Raise Err.Number, Err.Source, "DbRecordset->IDbRecordset_OpenRecordset->AdoRecordset.Open: " & Err.Description, Err.HelpFile, Err.HelpContext
End Function


Private Function IDbRecordset_OpenScalar(ByVal sql As String, ParamArray args() As Variant) As Variant
    '@Ignore FunctionReturnValueDiscarded
    IDbRecordset_AdoRecordset sql, args
    
    On Error GoTo RecordsetOpenError
    AdoRecordset.Open Options:=this.OpenOptions
    On Error GoTo 0
    If AdoRecordset.CursorLocation = adUseClient Then Set AdoRecordset.ActiveConnection = Nothing
    IDbRecordset_OpenScalar = AdoRecordset.Fields.Item(0).value
    Exit Function

RecordsetOpenError:
    Err.Raise Err.Number, Err.Source, "DbRecordset->IDbRecordset_OpenScalar->AdoRecordset.Open: " & Err.Description, Err.HelpFile, Err.HelpContext
End Function


' https://docs.microsoft.com/en-us/sql/ado/guide/data/ado-event-instantiation-visual-basic
' https://docs.microsoft.com/en-us/sql/ado/guide/data/ado-event-handler-summary
' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/fetchprogress-event-ado
'
'@Ignore ProcedureNotUsed, ParameterNotUsed, ParameterCanBeByVal
Private Sub AdoRecordset_FetchProgress(ByVal Progress As Long, ByVal MaxProgress As Long, ByRef adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Debug.Print "ADODB.Recordset fetch in progress. " & CStr(Progress) & " records of " & CStr(MaxProgress) & "have been retrived..."
End Sub


' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/fetchcomplete-event-ado
'
'@Ignore ProcedureNotUsed, ParameterNotUsed, ParameterCanBeByVal
Private Sub AdoRecordset_FetchComplete(ByVal pError As ADODB.Error, ByRef adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Debug.Print "ADODB.Recordset fetch completed. "
    If adStatus = ADODB.EventStatusEnum.adStatusErrorsOccurred And pError Is Nothing Then
        Debug.Print "Error: " & pError.Number, pError.Description, pError.Source, pError.Source
    End If
End Sub
