---
layout: default
title: DbRecordset
nav_order: 3
permalink: /dbrecordset
---

DbRecordset is a new class added to SecureADODB-PG, which wraps the ADODB.Recordset object. Functionally, DbRecordset is responsible for SELECT queries returning recordset objects, while DbCommand is still responsible for UPDATE/INSERT/DELETE queries, like in SecureADODB-RD.

DbRecordset includes members targeting several groups of tasks:

1. exposing class's private attributes
2. retrieving data from the database via SELECT queries
3. updating local structures and the database data
4. providing convenient development access to the recordset data.

### IDbRecordset interface

The IDbRecordset class formalizes the public interface of DbRecordset and exposes several methods and attributes.

**Attribute** members expose the two primary DbRecordset attributes, including instances of the ADODB.Recordset class and DbCommand/IDbCommad class:

```vb
Public Property Get cmd() As IDbCommand
End Property

Public Property Get AdoRecordset() As ADODB.Recordset
End Property

Public Function GetAdoRecordset(ByVal SQL As String, ParamArray ADODBParamsValues() As Variant) As Recordset
End Function
```

**SELECT** methods provide a means to execute queries returning either recordset or a scalar value:

```vb
Public Function OpenRecordset(ByVal SQL As String, ParamArray ADODBParamsValues() As Variant) As Recordset
End Function

Public Function OpenScalar(ByVal SQL As String, ParamArray ADODBParamsValues() As Variant) As Variant
End Function
```

**Update** methods provide a means to change the data in the recordset (UpdateRecord) and to persist changes (UpdateRecordset via updatable recordset):

```vb
Public Sub UpdateRecord(ByVal AbsolutePosition As Long, ByVal ValuesDict As Dictionary)
End Sub

Public Sub UpdateRecordset(ByRef AbsolutePositions() As Long, ByRef RecordsetData() As Variant)
End Sub
```

**Convenience** routines provide development access to recordset data. RecordsetToQT outputs recordset data onto an Excel worksheet via the QueryTable feature:

```vb
Public Function RecordsetToQT(ByVal OutputRange As Range) As QueryTable
End Function
```

### UpdateRecordset

A database can be updated via ADODB using either UPDATE/INSERT/DELETE SQL statements (typically with the Command object) or using updatable recordsets, both having their advantages and limitations. Here, I will focus on the latter option.

```vb
Private Sub IDbRecordset_UpdateRecordset(ByRef AbsolutePositions() As Long, ByRef RecordsetData() As Variant)
    UpdateRecordsetData AbsolutePositions, RecordsetData
    Dim DirtyRecordsCount As Long
    DirtyRecordsCount = UBound(AbsolutePositions) - LBound(AbsolutePositions) + 1
    PersistRecordsetChanges DirtyRecordsCount
End Sub
```

UpdateRecordset wraps the UpdateBatch method (ADODB.Recordset), and several factors affect a particular implementation of the additional wrapping code. The typical workflow involves an initial SELECT query retrieving data from the database into the recordset attribute of the DbRecordset class. Then the client modifies the data, and UpdateBatch can be used to persist the changes in the database. Data modification occurs outside of the library; therefore, the recordset data must be accessible to the user. Current implementation of SecureADODB-PG DbRecordset/IDbRecordset classes exposes the encapsulated ADODB.Recordset object, so the library user has two choices.

The user may copy recordset data to an independent local container, such as a 2D array. If there are any local changes, the data in the recordset needs to be updated first. Alternatively, the user may use the recordset object directly without an intermediate container, saving changes into the recordset as they occur. In either case, either the user or the SecureADODB library may perform the update process.

The first prospective user for this fork is the [ContactEditor][] demo app. In ContactEditor, SecrueADODB will interface with the [Storage Library][] employing the first strategy with a 2D array as an independent local data container. Two DbRecordset routines, *UpdateRecordsetData* and *PersistRecordsetChanges*, will update recordset data and the database, respectively.


<details><summary>Getters and Setters</summary>

```vb

Public Property Let FirstName(ByVal Value As String)
  FirstName = Value
End Property

Public Property Get FirstName() As String
  FirstName = this.FirstName
End Property

Public Property Let LastName(ByVal Value As String)
  LastName = Value
End Property

Public Property Get LastName() As String
  LastName = this.LastName
End Property

Public Property Let Login(ByVal Value As String)
  Login = Value
End Property

Public Property Get Login() As String
  Login = this.Login
End Property

```

</details>

___


<details><summary>DbRecordset.PersistRecordsetChanges</summary>

```vb

Friend Sub PersistRecordsetChanges(ByVal DirtyRecordCount As Long)
    With AdoRecordset
        Guard.ExpressionErr .State = adStateOpen, _
                            IncompatibleStatusErr, _
                            "DbRecordset", _
                            "Expected AdoRecordset.Status = adStateOpen"
                            
        Dim db As IDbConnection
        Set db = this.cmd.Connection
        '''' Marshal dirty records only
        .MarshalOptions = adMarshalModifiedOnly
        Set .ActiveConnection = this.cmd.Connection.AdoConnection
        On Error GoTo Rollback
        '''' Set the expected count of affected rows in the DbConnection object
        db.ExpectedRecordsAffected = DirtyRecordCount
        '''' Wrap update in a transaction
        db.BeginTransaction
        .UpdateBatch
        db.CommitTransaction
        On Error GoTo 0
        If .CursorLocation = adUseClient Then Set .ActiveConnection = Nothing
    End With
    
    Exit Sub
    
Rollback:
    this.cmd.Connection.RollbackTransaction
    With Err
        .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
    End With
End Sub

```

</details>

