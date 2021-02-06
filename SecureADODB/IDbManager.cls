VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Represents an object encapsulating a database transaction."
'@Folder "SecureADODB.DbManager.Abstract"
'@ModuleDescription("Represents an object encapsulating a database transaction.")
'@Interface
Option Explicit


'@Description("Returns class reference")
Public Property Get Class() As DbManager
Attribute Class.VB_Description = "Returns class reference"
End Property

'@Description("Creates a new command to execute as part of the transaction.")
Public Function Command() As IDbCommand
Attribute Command.VB_Description = "Creates a new command to execute as part of the transaction."
End Function

'@Description("Commits the transaction.")
Public Sub Commit()
Attribute Commit.VB_Description = "Commits the transaction."
End Sub

'@Description("Rolls back the transaction.")
Public Sub Rollback()
Attribute Rollback.VB_Description = "Rolls back the transaction."
End Sub
