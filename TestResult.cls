VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Implements TestEntry

Private pass As Boolean
Private number As Integer
Private testType As String
Private errorInfo As String
Private message As String

Public Sub init(p As Boolean, n As Integer, t As String, e As String, m As String)
    pass = p
    number = n
    testType = t
    errorInfo = e
    message = m
End Sub

Public Function errorCount() As Long
    errorCount = IIf(pass, 0, 1)
End Function

Public Function getError(Optional indentation As Integer = 0) As String
    Dim out As String
    out = ""

    If errorCount > 0 Then
        out = out & Stringx.repeat("  ", indentation)
        If message = "" Then
            out = out & Stringx.format("{0,3} {1} || {2}", number, testType, errorInfo)
        Else
            out = out & Stringx.format("{0,3} {1} || {2} || {3}", number, testType, errorInfo, message)
        End If
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

