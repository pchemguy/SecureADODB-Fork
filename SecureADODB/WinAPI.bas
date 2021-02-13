Attribute VB_Name = "WinAPI"
'@Folder("SecureADODB.Shared.WinAPI")
Option Explicit


#If VBA7 Then
    'For 64-Bit versions of Excel
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    'For 32-Bit versions of Excel
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
