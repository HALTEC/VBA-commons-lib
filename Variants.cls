VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Variants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Function isEquatable(item As Variant) As Boolean
    If IsObject(item) Then
        isEquatable = TypeOf item Is IEquatable
    Else
        isEquatable = True
    End If
End Function

Public Function isComparable(item As Variant) As Boolean
    If IsObject(item) Then
        isComparable = TypeOf item Is IComparable
    Else
        isComparable = True
    End If
End Function

Public Function equals(value As Variant, other As Variant) As Boolean
    If IsObject(value) <> IsObject(other) Then
        equals = False
    ElseIf IsObject(value) Then
        equals = value Is other
    Else
        equals = value = other
    End If
End Function

Public Function fuzzyEquals(value As Variant, other As Variant) As Boolean
    If Not isEquatable(value) Or Not isEquatable(other) Then
        Err.Raise E_ARGUMENTOUTOFRANGE, "Variants.Equals()", "Can only equate things that are Variants.IsEquatable()."
    End If
    
    If IsObject(value) <> IsObject(other) Then
        fuzzyEquals = False
    ElseIf IsObject(value) Then
    
        Dim equatable As IEquatable
        Set equatable = value
        fuzzyEquals = equatable.equals(other)
    
    Else
    
        fuzzyEquals = value = other
    
    End If
End Function

Public Function isArray(var As Variant) As Boolean
    ' TODO: Test
    isArray = CBool(varType(var) And vbArray)
End Function

Public Sub verifyArray(var As Variant)
    If Not isArray(var) Then
        Err.Raise E_ARGUMENTOUTOFRANGE, "", "The parameter has to be an array."
    End If
End Sub

Public Function gist(var As Variant) As String
    If varType(var) And vbArray Then
    
        gist = "Array<" & varTypeToString(varType(var) - vbArray) & ">"
        gist = gist & "["
        
        Dim inner As String
        Dim runner As Integer
        For runner = LBound(var) To UBound(var)
            If inner <> "" Then
                inner = inner & ", "
            End If
            inner = inner & gist(var(runner))
        Next
        gist = gist & inner
        
        gist = gist & "]"
    
    ElseIf varType(var) = vbBoolean _
        Or varType(var) = vbByte _
        Or varType(var) = vbCurrency _
        Or varType(var) = vbDate _
        Or varType(var) = vbDecimal _
        Or varType(var) = vbDouble _
        Or varType(var) = vbInteger _
        Or varType(var) = vbLong _
        Or varType(var) = vbSingle Then

        gist = CStr(var)
    
    ElseIf varType(var) = vbString Then
        
        gist = "'" & CStr(var) & "'"
    
    ElseIf varType(var) = vbDataObject Then
    
        gist = "DataObject"
    
    ElseIf varType(var) = vbEmpty Then
    
        gist = "Empty"
    
    ElseIf varType(var) = vbError Then
    
        gist = CStr(var) & " (Error)"
    
    ElseIf varType(var) = vbNull Then
    
        gist = "Null"
    
    ElseIf varType(var) = vbObject Then
    
        If var Is Nothing Then
            gist = TypeName(var)
        ElseIf TypeName(var) = "List" Then
            gist = var.gist
        Else
            gist = TypeName(var)
        End If
    
    ElseIf varType(var) = vbUserDefinedType Then
    
        gist = "UserDefinedType"
    
    ElseIf varType(var) = vbVariant Then
    
        gist = "Variant"
    
    End If
End Function

Public Function varTypeToString(varType As Long) As String
    If varType = vbBoolean Then
        varTypeToString = "Boolean"
    ElseIf varType = vbByte Then
        varTypeToString = "Byte"
    ElseIf varType = vbCurrency Then
        varTypeToString = "Currency"
    ElseIf varType = vbDataObject Then
        varTypeToString = "DataObject"
    ElseIf varType = vbDate Then
        varTypeToString = "Date"
    ElseIf varType = vbDecimal Then
        varTypeToString = "Decimal"
    ElseIf varType = vbDouble Then
        varTypeToString = "Double"
    ElseIf varType = vbEmpty Then
        varTypeToString = "Empty"
    ElseIf varType = vbInteger Then
        varTypeToString = "Integer"
    ElseIf varType = vbLong Then
        varTypeToString = "Long"
    ElseIf varType = vbNull Then
        varTypeToString = "Null"
    ElseIf varType = vbSingle Then
        varTypeToString = "Single"
    ElseIf varType = vbString Then
        varTypeToString = "String"
    ElseIf varType = vbObject Then
        varTypeToString = "Object"
    ElseIf varType = vbUserDefinedType Then
        varTypeToString = "UserDefinedType"
    ElseIf varType = vbVariant Then
        varTypeToString = "Variant"
    ElseIf varType = vbArray Then
        varTypeToString = "Array"
    Else
        varTypeToString = "Unknown: " & varType
    End If
End Function

