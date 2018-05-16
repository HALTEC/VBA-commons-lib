VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Math"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Function min(ParamArray values() As Variant) As Variant
    Dim minValue, value As Variant
    minValue = values(0)
    For Each value In values
        If value < minValue Then minValue = value
    Next
    min = minValue
End Function

Public Function max(ParamArray values() As Variant) As Variant
    Dim maxValue, value As Variant
    maxValue = values(0)
    For Each value In values
        If value > maxValue Then maxValue = value
    Next
    max = maxValue
End Function

Public Function cmp(value As Variant, other As Variant) As Integer
    If Not Variants.isComparable(value) Or Not Variants.isComparable(other) Then
        Err.Raise E_ARGUMENTOUTOFRANGE, "Math.cmp()", "Can only compare types that are Variants.isComparable()."
    End If
    
    If IsObject(value) <> IsObject(other) Then
        Err.Raise E_TYPEMISMATCH, "Math.cmp()", "Can only compare things of same type."
    End If
    
    If IsObject(value) Then
        Dim comparable As IComparable
        Set comparable = value
        cmp = comparable.compareTo(other)
    Else
        If value < other Then
            cmp = -1
        ElseIf value > other Then
            cmp = 1
        Else
            cmp = 0
        End If
    End If
End Function

Public Function ceiling(ByVal value As Double) As Integer
    If Int(value) = value Then
        ceiling = value
    Else
        ceiling = Int(value) + 1
    End If
End Function

Public Function floor(ByVal value As Double) As Integer
    floor = Int(value)
End Function
