---
layout: default
title: DbRecordset
nav_order: 2
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

Public Function GetAdoRecordset(ByVal SQL As String, _
								ParamArray ADODBParamsValues() As Variant) As ADODB.Recordset
End Function
```

**SELECT** methods provide a means to execute queries returning either recordset or a scalar value:


```vb
Public Function OpenRecordset(ByVal SQL As String, _
							  ParamArray ADODBParamsValues() As Variant) As ADODB.Recordset
End Function

Public Function OpenScalar(ByVal SQL As String, _
						   ParamArray ADODBParamsValues() As Variant) As Variant
End Function
```

**Update** methods provide a means to change the data in the recordset (UpdateRecord) and to persist changes (UpdateRecordset via updatable recordset):


```vb
Public Sub UpdateRecord(ByVal AbsolutePosition As Long, ByVal ValuesDict As Scripting.Dictionary)
End Sub

Public Sub UpdateRecordset(ByRef AbsolutePositions() As Long, ByRef RecordsetData() As Variant)
End Sub
```

**Convenience** routines provide development access to recordset data. RecordsetToQT outputs recordset data onto an Excel worksheet via the QueryTable feature:


```vb
Public Function RecordsetToQT(ByVal OutputRange As Excel.Range) As Excel.QueryTable
End Function
```

### UpdateRecordset

A database can be updated via ADODB using either UPDATE/INSERT/DELETE SQL statements (typically with the Command object) or using updatable recordsets, both having their advantages and limitations. Here, I will focus on the latter option.

UpdateRecordset wraps the UpdateBatch method (ADODB.Recordset), and several factors affect a particular implementation of the additional wrapping code. The typical workflow involves an initial SELECT query retrieving data from the database into the recordset attribute of the DbRecordset class. Then the client modifies the data, and UpdateBatch can be used to persist the changes in the database. Data modification occurs outside of the library; therefore, the recordset data must be accessible to the user. Current implementation of SecureADODB-PG DbRecordset/IDbRecordset classes exposes the encapsulated ADODB.Recordset object, so the library user has two choices.

The user may copy recordset data to an independent local container, such as a 2D array. If there are any local changes, the data in the recordset needs to be updated first. Alternatively, the user may use the recordset object directly without an intermediate container, saving changes into the recordset as they occur. In either case, either the user or the SecureADODB library may perform the update process.

The first prospective user for this fork is the [ContactEditor][] demo app. In ContactEditor, SecrueADODB will interface with the [Storage Library][] employing the first strategy with a 2D array as an independent local data container. Two DbRecordset routines, *UpdateRecordsetData* and *PersistRecordsetChanges*, will update recordset data and the database, respectively.

Apart from the invoking database update command, *PersistRecordsetChanges* incorporates two other features. It wraps the *UpdateBatch* call in a database transaction and verifies that the expected and actual number of changes match.

Some backends do not support transactions. In the current implementation, *PersistRecordsetChanges* raises an error when transactions are not available. It can handle this limitation more gracefully by checking the *TransactionsDisabled* flag (currently not exposed).

### Affected rows count

Verifying the affected rows count is a convenient and efficient consistency check. *UpdateRecordset* method takes a 1D array containing ids of dirty records. Therefore, the expected value for the number of affected rows is readily available. It appears, however, that the actual number is not available from the recordset object, necessitating the use of backend-specific sources.

In SQLite, `SELECT total_changes()` query returns the total number of changes for the Connection object used. If executed before and after the transaction wrapping the *UpdateBatch* call, it yields the number of rows changed by the database engine during the transaction. For it to work correctly, this query must share the Connection object with *UpdateBatch* and transaction-related commands. The first call (from the *BeginTransComplete* handler) caches the reference value in the *TransRecordsAffected* attribute (the *ExecuteComplete* handler sets a similar *RecordsAffected* variable). The second call (from the *CommitTransComplete*) yields the desired value and verifies that it matches the expected count.

Two additional DbConnection attributes (*Engine* and *cmdAffectedRows*) help streamline this engine-specific solution. *Engine* setter initializes both of these attributes when the DbManager.CreateFileDb factory sets *Engine* to its first argument, *DbType*. *cmdAffectedRows* is an ADODB.Command object set to retrieve the total changes count. Connection event handlers, in turn, call the *TotalChanges* function, which executes the *cmdAffectedRows* command and returns affected rows count or -1 if this feature is unavailable.



[ContactEditor]: https://pchemguy.github.io/ContactEditor/
[Storage Library]: https://pchemguy.github.io/ContactEditor/storage-library
