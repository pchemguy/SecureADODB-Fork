Attribute VB_Name = "CommonRoutines"
'@Folder "Common.Shared"
Option Explicit

'@IgnoreModule MoveFieldCloserToUsage
Private lastID As Double


'@EntryPoint
Public Function GetTimeStampMs() As String
    '''' On Windows, the Timer resolution is subsecond, the fractional part (the four characters at the end
    '''' given the format) is concatenated with DateTime. It appears that the Windows' high precision time
    '''' source available via API yields garbage for the fractional part.
    GetTimeStampMs = Format$(Now, "yyyy-MM-dd HH:mm:ss") & Right$(Format$(Timer, "#0.000"), 4)
End Function


'''' The number of seconds since the Epoch is multiplied by 10^4 to bring the first
'''' four fractional places in Timer value into the whole part before trancation.
'''' Long on a 32bit machine does not provide sufficient number of digits,
'''' so returning double. Alternatively, a Currency type could be used.
'@EntryPoint
Public Function GenerateSerialID() As Double
    Dim newID As Double
    Dim secTillLastMidnight As Double
    secTillLastMidnight = CDbl(DateDiff("s", DateSerial(1970, 1, 1), Date))
    newID = Fix((secTillLastMidnight + Timer) * 10 ^ 4)
    If newID > lastID Then
        lastID = newID
    Else
        lastID = lastID + 1
    End If
    GenerateSerialID = lastID
    'GetSerialID = Fix((CDbl(Date) * 100000# + CDbl(Timer) / 8.64))
End Function


'''' When sub/function captures a list of arguments in a ParamArray and passes it
'''' to the next routine expecting a list of arguments, the second routine receives
'''' a 2D array instead of 1D with the outer dimension having a single element.
'''' This function check the arguments and unfolds the outer dimesion as necessary.
'''' Any function accepting a ParamArray argument should be able to use it.
''''
'''' Unfold if the following conditions are satisfied:
''''     - ParamArrayArg is a 1D array
''''     - UBound(ParamArrayArg, 1) = LBound(ParamArrayArg, 1) = 0
''''     - ParamArrayArg(0) is a 1D 0-based array
''''
'''' Return
''''     - ParamArrayArg(0), if unfolding is necessary
''''     - ParamArrayArg, if ParamArrayArg is array, but not all conditions are satisfied
'''' Raise an error if is not an array
'@Description "Unfolds a ParamArray argument when passed from another ParamArray."
Public Function UnfoldParamArray(ByVal ParamArrayArg As Variant) As Variant
Attribute UnfoldParamArray.VB_Description = "Unfolds a ParamArray argument when passed from another ParamArray."
    Guard.NotArray ParamArrayArg
    Dim DoUnfold As Boolean
    DoUnfold = (ArrayLib.NumberOfArrayDimensions(ParamArrayArg) = 1) And (LBound(ParamArrayArg) = 0) And (UBound(ParamArrayArg) = 0)
    If DoUnfold Then DoUnfold = IsArray(ParamArrayArg(0))
    If DoUnfold Then DoUnfold = ((ArrayLib.NumberOfArrayDimensions(ParamArrayArg(0)) = 1) And (LBound(ParamArrayArg(0), 1) = 0))
    If DoUnfold Then
        UnfoldParamArray = ParamArrayArg(0)
    Else
        UnfoldParamArray = ParamArrayArg
    End If
End Function


'@EntryPoint
Public Function GetVarType(ByRef Variable As Variant) As String
    Dim NDim As String
    NDim = IIf(IsArray(Variable), "/Array", vbNullString)
    
    Dim TypeOfVar As VBA.VbVarType
    TypeOfVar = VarType(Variable) And Not vbArray

    Dim ScalarType As String
    Select Case TypeOfVar
        Case vbEmpty
            ScalarType = "vbEmpty"
        Case vbNull
            ScalarType = "vbNull"
        Case vbInteger
            ScalarType = "vbInteger"
        Case vbLong
            ScalarType = "vbLong"
        Case vbSingle
            ScalarType = "vbSingle"
        Case vbDouble
            ScalarType = "vbDouble"
        Case vbCurrency
            ScalarType = "vbCurrency"
        Case vbDate
            ScalarType = "vbDate"
        Case vbString
            ScalarType = "vbString"
        Case vbObject
            ScalarType = "vbObject"
        Case vbError
            ScalarType = "vbError"
        Case vbBoolean
            ScalarType = "vbBoolean"
        Case vbVariant
            ScalarType = "vbVariant"
        Case vbDataObject
            ScalarType = "vbDataObject"
        Case vbDecimal
            ScalarType = "vbDecimal"
        Case vbByte
            ScalarType = "vbByte"
        Case vbUserDefinedType
            ScalarType = "vbUserDefinedType"
        Case Else
            ScalarType = "vbUnknown"
    End Select
    GetVarType = ScalarType & NDim
End Function


'''' Resolves file pathname
''''
'''' This helper routines attempts to interpret provided pathname as
'''' a reference to an existing file:
'''' 1) check if provided reference is a valid absolute file pathname, if not,
'''' 2) check if provided reference is a valid name of file located in the same
''''    folder as this workbook (prefix Thisworkbook.path), if not,
'''' 3) throw FileNotFound Error if the second argument is empty
'''' 4) check if the second argument is a string; if so, cast it as a 1D array
''''    of size one and skip next
'''' 5) check if the second argument is an array of strings, if not,
''''    throw FileNotFound Error
'''' 6) check if provided reference is a valid absolute path containing file
''''    with name matching Thisworkbook.Name (both without extension) and
''''    extension taken from the second argument, if not
'''' 7) run the check above using Thisworkbook.Path as the target file location,
''''    if resolution fails, throw FileNotFound Error.
''''
'''' Args:
''''   FilePathName (string):
''''     File pathname
''''
''''   DefaultExts (string or string/array):
''''     1D array of default extensions or a single default extension
''''
'''' Returns:
''''   String:
''''     Resolved valid absolute pathname pointing to an existing file.
''''
'''' Throws:
''''   Err.FileNotFoundErr:
''''     If provided pathname cannot be resolved to a valid file pathname.
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
'@Description "Resolves file pathname"
Public Function VerifyOrGetDefaultPath(ByVal FilePathName As String, ByVal DefaultExts As Variant) As String
Attribute VerifyOrGetDefaultPath.VB_Description = "Resolves file pathname"
    '''' Match any file with Dir$ igonring attributes
    Const vbAnyAttr As Long = vbNormal + vbReadOnly + vbHidden + vbSystem + vbArchive
    
    Dim FileExist As Variant
    Dim PathNameCandidate As String
    
    If Len(FilePathName) > 0 Then
        '''' If matched, Dir returns Len(String) > 0;
        '''' otherwise, returns vbNullString or raises an error
        
        '''' === (1) === Check if FilePathName is a valid path to an existing file.
        PathNameCandidate = FilePathName
        On Error Resume Next
        FileExist = Dir$(PathNameCandidate, vbAnyAttr)
        On Error GoTo 0
        If Len(FileExist) > 0 Then
            VerifyOrGetDefaultPath = PathNameCandidate
            Exit Function
        End If
        
        '''' === (2) === Check if FilePathName is a valid name of file located in
        ''''             the same folder as this workbook
        PathNameCandidate = ThisWorkbook.Path & Application.PathSeparator & FilePathName
        On Error Resume Next
        FileExist = Dir$(PathNameCandidate, vbAnyAttr)
        On Error GoTo 0
        If Len(FileExist) > 0 Then
            '''' Return resolved absolute file pathname
            VerifyOrGetDefaultPath = PathNameCandidate
            Exit Function
        End If
    End If
    
    '''' === (3) === Check if the second argument is valid
    Dim ValidExt As Variant
    On Error Resume Next
    If VarType(DefaultExts) >= vbArray Then
        ValidExt = (VarType(DefaultExts(0)) = vbString)
    Else
        ValidExt = (VarType(DefaultExts) = vbString)
    End If
    On Error GoTo 0
    
    If Not ValidExt Then
        VBA.Err.Raise _
            Number:=ErrNo.FileNotFoundErr, _
            Source:="DataTableADODB", _
            Description:="File <" & FilePathName & "> not found!"
    End If
        
    Dim Extensions As Variant
    If VarType(DefaultExts) = vbString Then
        Extensions = Array(DefaultExts)
    Else
        Extensions = DefaultExts
    End If
    
    Dim NameRoot As String
    NameRoot = ThisWorkbook.Name
    Dim DotPos As Long
    DotPos = InStr(Len(NameRoot) - 5, NameRoot, ".xl", vbTextCompare)
    NameRoot = Left$(NameRoot, DotPos)
    
    Dim Prefix As String
    Dim ExtIndex As Long
    
    '''' === (6) === FilePathName is a valid path
    On Error Resume Next
    Prefix = Dir$(FilePathName, vbAnyAttr)
    On Error GoTo 0
    
    If Len(Prefix) > 0 Then
        Prefix = FilePathName & Application.PathSeparator
        For ExtIndex = LBound(Extensions) To UBound(Extensions)
            PathNameCandidate = Prefix & NameRoot & Extensions(ExtIndex)
            FileExist = Dir$(PathNameCandidate)
            If Len(FileExist) > 0 Then
                VerifyOrGetDefaultPath = PathNameCandidate
                Exit Function
            End If
        Next ExtIndex
    End If
    
    '''' === (7) === Use extensions only
    Prefix = ThisWorkbook.Path & Application.PathSeparator
    For ExtIndex = LBound(Extensions) To UBound(Extensions)
        PathNameCandidate = Prefix & NameRoot & Extensions(ExtIndex)
        FileExist = Dir$(PathNameCandidate)
        If Len(FileExist) > 0 Then
            VerifyOrGetDefaultPath = PathNameCandidate
            Exit Function
        End If
    Next ExtIndex
    
    VBA.Err.Raise _
        Number:=ErrNo.FileNotFoundErr, _
        Source:="DataTableADODB", _
        Description:="File <" & FilePathName & "> not found!"
End Function
