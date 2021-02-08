VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoParameterProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for converting a value of a mapped data type to an ADODB.Parameter object."
'@Folder "SecureADODB.DbParameterProvider"
'@ModuleDescription("An object responsible for converting a value of a mapped data type to an ADODB.Parameter object.")
'@PredeclaredId
'@IgnoreModule UseMeaningfulName, IllegalAnnotation
'@Exposed
Option Explicit

Implements IParameterProvider

Private mappings As ITypeMap


'@Inject: just an idea.. see #https://github.com/rubberduck-vba/Rubberduck/issues/5463
'@Description("Creates a new parameter provider using the specified type mappings.")
Public Function Create(ByVal typeMap As ITypeMap) As IParameterProvider
Attribute Create.VB_Description = "Creates a new parameter provider using the specified type mappings."
    Errors.GuardNonDefaultInstance Me, AdoParameterProvider
    
    Dim result As AdoParameterProvider
    Set result = New AdoParameterProvider
    result.CreInit typeMap
    Set Create = result
End Function


'@Description("Default constructor")
Friend Sub CreInit(ByVal typeMap As ITypeMap)
Attribute CreInit.VB_Description = "Default constructor"
    Errors.GuardNullReference typeMap
    Set mappings = typeMap
End Sub


Private Function IParameterProvider_FromValues(ByRef values() As Variant) As VBA.Collection
    Dim results As VBA.Collection
    Set results = New VBA.Collection
    Dim i As Long
    For i = LBound(values) To UBound(values)
        results.Add IParameterProvider_FromValue(values(i))
    Next
    Set IParameterProvider_FromValues = results
End Function


Private Function IParameterProvider_FromValue(ByVal value As Variant, Optional ByVal Name As String, Optional ByVal dataTypeName As String) As ADODB.Parameter
    If TypeOf value Is ADODB.Parameter Then
        Set IParameterProvider_FromValue = value
        Exit Function
    End If

    Dim localDataTypeName As String
    localDataTypeName = dataTypeName
    If localDataTypeName = vbNullString Then localDataTypeName = TypeName(value)
    Errors.GuardExpression Not mappings.IsMapped(localDataTypeName), message:="The data type '" & localDataTypeName & "' has no ADODB.DataTypeEnum mapping."
    
    Dim result As ADODB.Parameter
    Set result = New ADODB.Parameter
    result.Type = mappings.Mapping(localDataTypeName)
    result.direction = ADODB.ParameterDirectionEnum.adParamInput
    
    If result.Type = AdoTypeMappings.DefaultNullMapping Then
        result.Attributes = ADODB.ParameterAttributesEnum.adParamNullable
        result.value = Null
    Else
        result.value = value
    End If
    
    If IsCharMapping(result.Type) And Len(value) > 0 Then result.Size = Len(value)
    If Name <> vbNullString Then result.Name = Name
    
    Set IParameterProvider_FromValue = result
End Function


Private Function IsCharMapping(ByVal value As ADODB.DataTypeEnum) As Boolean
    IsCharMapping = ((value = adChar) _
                  Or (value = adWChar) _
                  Or (value = adVarChar) _
                  Or (value = adVarWChar))
End Function
