VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Represents an object that wraps an ADODB recordset."
'@Folder "SecureADODB.DbRecordset"
'@ModuleDescription("Represents an object that wraps an ADODB recordset.")
'@Interface
Option Explicit


Public Property Get Self() As IDbRecordset
End Property

'@Description("Gets the wrapped ADODB recordset.")
Public Property Get AdoRecordset() As ADODB.Recordset
Attribute AdoRecordset.VB_Description = "Gets the wrapped ADODB recordset."
End Property


'Public Sub OpenRecordset()
'End Sub
