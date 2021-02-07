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


'@Description "Adds item to the log database"
Public Sub Log(ByVal LogLine As String)
Attribute Log.VB_Description = "Adds item to the log database"
End Sub

'@Description "Prints the log database to the immediate"
Public Sub PrintLog()
Attribute PrintLog.VB_Description = "Prints the log database to the immediate"
End Sub

'@Description("Returns instance reference")
Public Property Get Self() As ILogger
Attribute Self.VB_Description = "Returns instance reference"
End Property

'@Description("Returns class reference")
Public Property Get Class() As ILogger
Attribute Class.VB_Description = "Returns class reference"
End Property
