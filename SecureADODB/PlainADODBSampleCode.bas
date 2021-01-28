Attribute VB_Name = "PlainADODBSampleCode"
'@Folder("-- DraftsTemplatesSnippets --")
'@IgnoreModule
Option Explicit


Public Sub TestADODBSourceSQL()
    Dim sDriver As String
    Dim sOptions As String
    Dim sDatabase As String

    Dim adoConnStr As String
    Dim qtConnStr As String
    Dim sSQL As String
    Dim sQTName As String
    
    sDatabase = ThisWorkbook.Path + "\" + "SecureADODB.db"
    sDriver = "{SQLite3 ODBC Driver}"
    sOptions = "SyncPragma=NORMAL;LongNames=True;NoCreat=True;FKSupport=True;OEMCP=True;"
    adoConnStr = "Driver=" + sDriver + ";" + _
                 "Database=" + sDatabase + ";" + _
                 sOptions

    qtConnStr = "OLEDB;" + adoConnStr
    
    sSQL = "SELECT * FROM categories WHERE category_id <= 3 AND section = 'machinery'"
    
    Dim adoRecordSet As ADODB.Recordset
    Set adoRecordSet = New ADODB.Recordset
    adoRecordSet.CursorLocation = adUseClient
    adoRecordSet.Open Source:=sSQL, ActiveConnection:=adoConnStr, CursorType:=adOpenKeyset, LockType:=adLockReadOnly, Options:=(adCmdText Or adAsyncFetch)
    Set adoRecordSet.ActiveConnection = Nothing
End Sub


Public Sub TestADODBSourceCMD()
    Dim sDriver As String
    Dim sOptions As String
    Dim sDatabase As String

    Dim adoConnStr As String
    Dim qtConnStr As String
    Dim sSQL As String
    Dim sQTName As String
    
    sDatabase = ThisWorkbook.Path + "\" + "SecureADODB.db"
    sDriver = "{SQLite3 ODBC Driver}"
    sOptions = "SyncPragma=NORMAL;LongNames=True;NoCreat=True;FKSupport=True;OEMCP=True;"
    adoConnStr = "Driver=" + sDriver + ";" + _
                 "Database=" + sDatabase + ";" + _
                 sOptions

    qtConnStr = "OLEDB;" + adoConnStr
    
    sSQL = "SELECT * FROM categories WHERE category_id <= 3 AND section = 'machinery'"
    
    Dim adoRecordSet As ADODB.Recordset
    Set adoRecordSet = New ADODB.Recordset
    Dim adoCommand As ADODB.Command
    Set adoCommand = New ADODB.Command
    
    With adoCommand
        .commandType = adCmdText
        .CommandText = sSQL
        .ActiveConnection = adoConnStr
        .ActiveConnection.CursorLocation = adUseClient
    End With
    
    With adoRecordSet
        Set .Source = adoCommand
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .Open Options:=adAsyncFetch
        Set .ActiveConnection = Nothing
    End With
    adoCommand.ActiveConnection.Close
End Sub


' Could not make it to work with named parameters
Public Sub TestADODBSourceCMDwithParametersPositional()
    Dim sDriver As String
    Dim sOptions As String
    Dim sDatabase As String

    Dim adoConnStr As String
    Dim qtConnStr As String
    Dim sSQL As String
    Dim sQTName As String
    
    sDatabase = ThisWorkbook.Path + "\" + "SecureADODB.db"
    sDriver = "{SQLite3 ODBC Driver}"
    sOptions = "SyncPragma=NORMAL;LongNames=True;NoCreat=True;FKSupport=True;OEMCP=True;"
    adoConnStr = "Driver=" + sDriver + ";" + _
                 "Database=" + sDatabase + ";" + _
                 sOptions

    qtConnStr = "OLEDB;" + adoConnStr
    
    sSQL = "SELECT * FROM categories WHERE category_id <= ? AND section = ?"
    
    Dim adoRecordSet As ADODB.Recordset
    Set adoRecordSet = New ADODB.Recordset
    Dim adoCommand As ADODB.Command
    Set adoCommand = New ADODB.Command
    
    Dim mappings As ITypeMap
    Set mappings = AdoTypeMappings.Default
    Dim provider As IParameterProvider
    Set provider = AdoParameterProvider.Create(mappings)
    
    Dim adoParameter As ADODB.Parameter
    Set adoParameter = provider.FromValue(3)
    'adoParameter.name = "@category_id"
    adoCommand.Parameters.Append adoParameter
    Set adoParameter = provider.FromValue("machinery")
    'adoParameter.name = "@section"
    adoCommand.Parameters.Append adoParameter
    
    With adoCommand
        .commandType = adCmdText
        .CommandText = sSQL
        .Prepared = True
        '.NamedParameters = True
        .ActiveConnection = adoConnStr
        .ActiveConnection.CursorLocation = adUseClient
    End With
        
    With adoRecordSet
        Set .Source = adoCommand
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .Open Options:=adAsyncFetch
        Set .ActiveConnection = Nothing
    End With
    adoCommand.ActiveConnection.Close
    Debug.Print "RecordCount: " & CStr(adoRecordSet.RecordCount)
End Sub


