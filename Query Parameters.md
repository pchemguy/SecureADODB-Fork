---
layout: default
title: Query parameters
nav_order: 2
permalink: /query-parameters
---

The AdoParameterProvider and AdoTypeMappings classes ([SecureADODB-RD][])  assist with parameterized queries. SecureADODB-PG modifies AdoTypeMappings and replaces AdoParameterProvider with the DbParameters class.

### AdoTypeMappings

AdoTypeMappings wraps the Scripting.Dictionary class and provides the mapping from standard VBA types to ADODB.DataTypeEnum. The present fork incorporates several small changes to this class.

First, the default mapping \<String&nbsp;&rarr;&nbsp;adVar**W**Char\> provided by AdoTypeMappings is not suitable for the CSV backend (as observed on Windows 10, MS Excel 2002-x32, using the stock Microsoft Text driver). Apparently, it expects the \<String&nbsp;&rarr;&nbsp;adVarChar\> mapping instead. Hence, a second factory, AdoTypeMappings.CSV, provides mapping suitable for the CSV backend. The DbManager.CreateFileDb factory takes database type as its first argument and uses it to select appropriate type mapping automatically.

Second, \<Null/Empty&nbsp;&rarr;&nbsp;adEmpty\> mapping did not work for me either. While the Type attribute of a standalone ADODB.Parameter accepts adEmpty value, an attempt to append such a parameter to the ADODB.Command.Parameters collection causes "inconsistent Parameter settings" error. I could not resolve this issue, so I switched to \<Null/Empty&nbsp;&rarr;&nbsp;adVarChar\> mapping with Variant/Null value.

Finally, a predeclared class should not, in general, contain Class_Initialize. If present, this routine is also executed during the initialization of the predeclared instance, and аny instructions relevant for non-default instances only will waste resources and complicate debugging executing unnecessary instructions. For this reason, AdoTypeMappings.Class_Initialize is replaced with *InitDefault* and *InitCSV* constructors.

### DbParameters

In ADODB, the Command class is responsible for handling parameterized queries. In particular, its CommandText attribute is set to an SQL query containing value placeholders, and its Parameters attribute is populated with ADODB.Parameter objects, one for each placeholder.

AdoParameterProvider (SecureADODB-RD) acts as an abstract factory for the ADODB.Parameter class (*generation*). Additionally, basic parameter *validation* is performed in the DbCommandBase class, and the calling class is responsible for *populating* the Parameters collection (DbCommandBase, acts as an abstract factory for the ADODB.Command class and populates Parameters).

DbParameters (SecureADODB-PG) is responsible for all three stages, *validation*, *generation*, and *population*. The IDbParameters class formalizes its public interface and provides one procedure, *FromValues*, which takes ADODB.Command and a parameter value list. *FromValues* performs several consistency checks: if the CommandText attribute is not blank, the counts of value and placeholder should match; if the Parameters collection is not empty, the numbers of existing parameters and values should match. Then the *FromValues* routine loops through the provided parameter value list. The routine either updates existing or creates new Parameter objects following the logic of AdoParameterProvider.FromValue appending the new Parameter objects to the Parameters collection.

The other IDbParameters routine, *GetSQL*, returns an interpolated SQL query. It quotes textual values, escaping single quotes, if necessary.


<!-- References -->

[SecureADODB-RD]: https://rubberduckvba.wordpress.com/2020/04/22/secure-adodb/
