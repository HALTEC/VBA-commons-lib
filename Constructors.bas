Attribute VB_Name = "Constructors"
Option Explicit
Option Base 0 ' Default


Public Function List_create(ParamArray values() As Variant) As List
' Create an untyped list.
    Dim arr As Variant: arr = values
    
    Dim l As New List
    l.init arr
    
    Set List_create = l
End Function

Public Function List_createT(typeString As String) As List
' Create a typed list.
    Dim l As New List
    l.initT typeString
    
    Set List_createT = l
End Function

Public Function List_createLT(ParamArray values() As Variant) As List
' Create a lazily typed list.
    Dim arr As Variant
    arr = values
    
    Dim l As New List
    l.initLT arr
    
    Set List_createLT = l
End Function

Public Function Map_create(ParamArray values() As Variant) As Map
    Dim arr As Variant: arr = values
    Dim m As New Map
    m.init arr

    Set Map_create = m
End Function

Public Function Map_createT(keyType As String, valueType As String) As Map
    Dim m As New Map
    m.initT keyType, valueType
    
    Set Map_createT = m
End Function

Public Function Map_createLT(ParamArray values() As Variant) As Map
    Dim arr As Variant: arr = values
    Dim m As New Map
    m.initLT arr
    
    Set Map_createLT = m
End Function

Public Function TestResult_create(pass As Boolean, number As Integer, testType As String, errorInfo As String, message As String) As TestResult
    Dim t As New TestResult
    t.init pass, number, testType, errorInfo, message
    Set TestResult_create = t
End Function

Public Function TestGroup_create(name As String) As TestGroup
    Dim t As New TestGroup
    t.init name
    Set TestGroup_create = t
End Function
