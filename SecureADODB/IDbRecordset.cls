VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object that wraps an ADODB recordset."
'@IgnoreModule FunctionReturnValueDiscarded
'@Folder "SecureADODB.DbRecordset"
'@ModuleDescription("Represents an object that wraps an ADODB recordset.")
'@Interface
'@Exposed
Option Explicit


Public Property Get Self() As IDbRecordset
End Property

'@Description("Gets the wrapped ADODB recordset.")
Public Property Get AdoRecordset() As ADODB.Recordset
Attribute AdoRecordset.VB_Description = "Gets the wrapped ADODB recordset."
End Property

Public Sub OpenRecordset(ByVal sql As String, ParamArray args() As Variant)
End Sub

Public Function OpenScalar(ByVal sql As String, ParamArray args() As Variant) As Variant
End Function
