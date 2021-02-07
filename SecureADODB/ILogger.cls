VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SecureADODB.Logger")
'@Interface
Option Explicit


'@Description("Returns instance reference")
Public Property Get Self() As ILogger
End Property

'@Description("Returns class reference")
Public Property Get Class() As ILogger
End Property

'@Description("Returns instance reference")
Public Property Get LogDatabase() As Scripting.Dictionary
End Property


'@Description "Adds item to the log database"
Public Sub Log(ByVal LogLine As String, Optional ByRef LogDatabase As Scripting.Dictionary = Nothing)
Attribute Log.VB_Description = "Adds item to the log database"
End Sub

'@Description "Prints the log database to the immediate"
Public Sub PrintLog(Optional ByRef LogDatabase As Scripting.Dictionary = Nothing)
Attribute PrintLog.VB_Description = "Prints the log database to the immediate"
End Sub

'@Description "Resets the log database"
Public Sub ClearLog(Optional ByRef LogDatabase As Scripting.Dictionary = Nothing)
End Sub

