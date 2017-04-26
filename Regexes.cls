VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Regexes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0 ' default

Public globalx As Integer
Public ignoreCase As Integer

Public Function replace(ByVal text As String, ByVal pattern As String, ByVal replacement As String, Optional ByVal options As Integer = 0) As String
    Dim regObj As New RegExp
    
    regObj.Global = options And globalx <> 0
    regObj.ignoreCase = options And ignoreCase <> 0
    
    regObj.pattern = pattern
    replace = regObj.replace(text, replacement)
End Function

Function escape(ByVal regex As String) As String
    regex = VBA.Strings.replace(regex, "\", "\\")
    regex = VBA.Strings.replace(regex, "^", "\^")
    regex = VBA.Strings.replace(regex, "$", "\$")
    regex = VBA.Strings.replace(regex, "?", "\?")
    regex = VBA.Strings.replace(regex, "*", "\*")
    regex = VBA.Strings.replace(regex, "+", "\+")
    regex = VBA.Strings.replace(regex, ".", "\.")
    regex = VBA.Strings.replace(regex, "|", "\|")
    regex = VBA.Strings.replace(regex, "{", "\{")
    regex = VBA.Strings.replace(regex, "}", "\}")
    regex = VBA.Strings.replace(regex, "[", "\[")
    regex = VBA.Strings.replace(regex, "]", "\]")
    regex = VBA.Strings.replace(regex, "(", "\(")
    regex = VBA.Strings.replace(regex, ")", "\)")
    
    escape = regex
End Function

Function escapeReplacement(ByVal replacement As String) As String
    escapeReplacement = VBA.Strings.replace(replacement, "$", "$$")
End Function

Private Sub Class_Initialize()
    globalx = 1
    ignoreCase = 2
End Sub
