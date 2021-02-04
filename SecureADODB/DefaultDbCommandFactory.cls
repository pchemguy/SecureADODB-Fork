VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DefaultDbCommandFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "SecureADODB.DbCommand"
'@IgnoreModule UseMeaningfulName, IllegalAnnotation
'@PredeclaredId
'@Exposed
Option Explicit

Implements IDbCommandFactory

Private Type TFactory
    Base As IDbCommandBase
End Type
Private this As TFactory


'@Inject
Public Function Create(ByVal baseCommand As IDbCommandBase) As IDbCommandFactory
    Errors.GuardNonDefaultInstance Me, DbCommandFactory
    
    Dim result As DefaultDbCommandFactory
    Set result = New DefaultDbCommandFactory
    result.CreInit baseCommand
    Set Create = result
End Function


'@Description("Default constructor")
Friend Sub CreInit(ByVal baseCommand As IDbCommandBase)
Attribute CreInit.VB_Description = "Default constructor"
    Errors.GuardNullReference baseCommand
    Set this.Base = baseCommand
End Sub


Private Function IDbCommandFactory_CreateInstance(ByVal db As IDbConnection) As IDbCommand
    Set IDbCommandFactory_CreateInstance = DefaultDbCommand.Create(db, this.Base)
End Function
