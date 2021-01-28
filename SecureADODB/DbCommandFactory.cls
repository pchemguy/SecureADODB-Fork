VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbCommandFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An implementation of an abstract factory that creates DbCommand objects."
'@Folder("SecureADODBmod.DbCommand")
'@IgnoreModule UseMeaningfulName, IllegalAnnotation
'@PredeclaredId
'@ModuleDescription("An implementation of an abstract factory that creates DbCommand objects.")

Option Explicit


Implements IDbCommandFactory


Private Type TFactory
    Base As IDbCommandBase
    commandClass As Object
End Type
Private this As TFactory


'@Description("Default abstract factory")
'@Ignore ProcedureNotUsed
Public Function CreateFactory(ByVal baseCommand As IDbCommandBase, Optional ByVal commandClass As Object = Nothing) As IDbCommandFactory
Attribute CreateFactory.VB_Description = "Default abstract factory"
    Dim result As DbCommandFactory
    Set result = New DbCommandFactory
    result.init baseCommand, commandClass
    Set CreateFactory = result
End Function


'@Description("Default constructor")
Public Sub init(ByVal baseCommand As IDbCommandBase, Optional ByVal commandClass As Object = Nothing)
Attribute init.VB_Description = "Default constructor"
    Set this.Base = baseCommand
    If commandClass Is Nothing Then
        Set this.commandClass = DefaultDbCommand
    Else
        Set this.commandClass = commandClass
    End If
End Sub


'@Description("Default interface factory")
Private Function IDbCommandFactory_CreateInstance(ByVal db As IDbConnection) As IDbCommand
Attribute IDbCommandFactory_CreateInstance.VB_Description = "Default interface factory"
    Set IDbCommandFactory_CreateInstance = this.commandClass.Create(db, this.Base)
End Function
