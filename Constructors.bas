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

Public Function Escaper_create(esc As String, quote As String, separators As List) As Escaper
    Dim escObj As New Escaper
    escObj.init esc, quote, separators
    Set Escaper_create = escObj
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

Public Function InputValue_create(name As String, value As String, Optional typex As String = "")
    Dim i As New InputValue
    i.init name, value, typex
    Set InputValue_create = i
End Function
