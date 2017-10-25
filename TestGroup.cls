VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Implements TestEntry

Private name As String
Private subTests As List
Private testCounter As Integer

Public Sub init(n As String)
    name = n
    Set subTests = List_create
    testCounter = 0
End Sub


Public Sub addTest(pass As Boolean, testType As String, errorInfo As String, message As String)
    testCounter = testCounter + 1
    Dim t As TestResult: Set t = TestResult_create(pass, testCounter, testType, errorInfo, message)
    
    subTests.push t
End Sub

Public Sub addGroup(group As TestGroup)
    subTests.push group
End Sub

Public Function errorCount() As Long
    Dim count As Long: count = 0
    
    Dim subTest As Variant
    For Each subTest In subTests
        count = count + subTest.errorCount
    Next

    errorCount = count
End Function

Public Function getError(Optional indentation As Integer = 0) As String
    Dim out As String
    out = ""

    If errorCount > 0 Then
        out = out & Stringx.repeat("  ", indentation)
        out = out & name & vbCrLf

        Dim subTest As Variant
        For Each subTest In subTests
            If subTest.errorCount > 0 Then
                out = out & subTest.getError(indentation + 1)
                out = out & vbCrLf
            End If
        Next
    End If

    getError = out
End Function


' Interface TestEntry

Private Function TestEntry_errorCount() As Long
    TestEntry_errorCount = errorCount
End Function

Private Function TestEntry_getError(Optional indentation As Integer = 0) As String
    TestEntry_getError = getError(indentation)
End Function
