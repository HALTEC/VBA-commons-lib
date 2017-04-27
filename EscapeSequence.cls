VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EscapeSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type tEscapeSequence
    EscapeString As String
    ReplacementString As String
End Type

Private this As tEscapeSequence
Option Explicit

Public Property Get EscapeString() As String
    EscapeString = this.EscapeString
End Property

Friend Property Let EscapeString(value As String)
    this.EscapeString = value
End Property

Public Property Get ReplacementString() As String
    ReplacementString = this.ReplacementString
End Property

Friend Property Let ReplacementString(value As String)
    this.ReplacementString = value
End Property

Public Function Create(escape As String, replacement As String) As EscapeSequence
    Dim result As New EscapeSequence
    result.EscapeString = escape
    result.ReplacementString = replacement
    Set Create = result
End Function
