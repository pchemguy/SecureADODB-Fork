Attribute VB_Name = "TypeMappingTests"
'@Folder "SecureADODB.DbParameterProvider.Tests"
'@TestModule
'@IgnoreModule
Option Explicit
Option Private Module

Private Const InvalidTypeName As String = "this isn't a valid type name"

#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If


'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub


'@TestMethod("Factory Guard")
Private Sub Default_ThrowsIfNotInvokedFromDefaultInstance()
    On Error GoTo TestFail
    With New AdoTypeMappings
        On Error GoTo CleanFail
        Dim sut As AdoTypeMappings
        Set sut = .Default
        On Error GoTo 0
    End With
CleanFail:
    If Err.Number = ErrNo.NonDefaultInstanceErr Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub


Private Sub DefaultMapping_MapsType(ByVal Name As String)
    Dim sut As ITypeMap
    Set sut = AdoTypeMappings.Default
    Assert.IsTrue sut.IsMapped(Name)
End Sub


'@TestMethod("Type Mappings")
Private Sub Mapping_ThrowsIfUndefined()
    On Error GoTo TestFail
    With AdoTypeMappings.Default
        On Error GoTo CleanFail
        Dim value As ADODB.DataTypeEnum
        value = .Mapping(InvalidTypeName)
        On Error GoTo 0
    End With
CleanFail:
    If Err.Number = ErrNo.CustomErr Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("Type Mappings")
Private Sub IsMapped_FalseIfUndefined()
    Dim sut As ITypeMap
    Set sut = AdoTypeMappings.Default
    Assert.IsFalse sut.IsMapped(InvalidTypeName)
End Sub


'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsBoolean()
    Dim value As Boolean
    DefaultMapping_MapsType TypeName(value)
End Sub


'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsByte()
    Dim value As Byte
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsCurrency()
    Dim value As Currency
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsDate()
    Dim value As Date
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsDouble()
    Dim value As Double
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsInteger()
    Dim value As Integer
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsLong()
    Dim value As Long
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsSingle()
    Dim value As Single
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsString()
    Dim value As String
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsEmpty()
    Dim value As Variant
    DefaultMapping_MapsType TypeName(value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsNull()
    Dim value As Variant
    value = Null
    DefaultMapping_MapsType TypeName(value)
End Sub

Private Function GetDefaultMappingFor(ByVal Name As String) As ADODB.DataTypeEnum
    On Error GoTo CleanFail
    Dim sut As ITypeMap
    Set sut = AdoTypeMappings.Default
    GetDefaultMappingFor = sut.Mapping(Name)
    Exit Function
CleanFail:
    Assert.Inconclusive "Default mapping is undefined for '" & Name & "'."
End Function

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForBoolean_MapsTo_adBoolean()
    Const Expected = adBoolean
    Dim value As Boolean
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForByte_MapsTo_adInteger()
    Const Expected = adInteger
    Dim value As Byte
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForCurrency_MapsTo_adCurrency()
    Const Expected = adCurrency
    Dim value As Currency
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForDate_MapsTo_adDate()
    Const Expected = adDate
    Dim value As Date
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForDouble_MapsTo_adDouble()
    Const Expected = adDouble
    Dim value As Double
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForInteger_MapsTo_adInteger()
    Const Expected = adInteger
    Dim value As Integer
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForLong_MapsTo_adInteger()
    Const Expected = adInteger
    Dim value As Long
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForNull_MapsTo_DefaultNullMapping()
    Dim Expected As ADODB.DataTypeEnum
    Expected = AdoTypeMappings.DefaultNullMapping
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(Null))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForEmpty_MapsTo_DefaultNullMapping()
    Dim Expected As ADODB.DataTypeEnum
    Expected = AdoTypeMappings.DefaultNullMapping
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(Empty))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForSingle_MapsTo_adSingle()
    Const Expected = adSingle
    Dim value As Single
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForString_MapsTo_adVarWChar()
    Const Expected = adVarWChar
    Dim value As String
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(value))
End Sub

