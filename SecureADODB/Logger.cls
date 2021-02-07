VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Logging facility using a dictionary object as the database."
'@Folder("SecureADODB.Logger")
'@ModuleDescription "Logging facility using a dictionary object as the database."
'@PredeclaredId
Option Explicit

Implements ILogger

Private Type TLogger
    Database As Scripting.Dictionary
End Type
Private this As TLogger


Private Sub Class_Initialize()
    Set this.Database = New Scripting.Dictionary
    this.Database.CompareMode = TextCompare
End Sub

Private Sub Class_Terminate()
    Set this.Database = Nothing
End Sub


'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As Logger
Attribute Class.VB_Description = "Returns class reference"
    Set Class = Logger
End Property


'@Description("Default class factory")
Public Function Create() As ILogger
Attribute Create.VB_Description = "Default class factory"
    Guard.NonDefaultInstance Me
    
    Dim result As Logger
    Set result = New Logger
    Set Create = result
End Function


Private Function GetTimeStampMs() As String
    GetTimeStampMs = Strings.Format(Now, "yyyy-MM-dd HH:mm:ss") & Strings.Right(Strings.Format(Timer, "#0.000"), 4)
End Function


'===================================================='
'================= ILogger INTERFACE ================'
'===================================================='

Private Property Get ILogger_Self() As ILogger
    Set ILogger_Self = Me
End Property


Private Property Get ILogger_Class() As ILogger
    Set ILogger_Class = Class
End Property


Private Sub ILogger_Log(ByVal LogLine As String)
    Dim key As String
    Do
        key = GetTimeStampMs
    Loop While this.Database.Exists(key)
    this.Database.Item(GetTimeStampMs) = LogLine
End Sub


Private Sub ILogger_PrintLog()
    Dim key As Variant
    For Each key In this.Database.Keys
        Debug.Print key & vbTab & vbTab & this.Database.Item(key)
    Next key
End Sub
