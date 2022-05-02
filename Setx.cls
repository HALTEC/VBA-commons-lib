VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Setx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private d As Scripting.Dictionary
Private isTyped As Boolean
Private innerType As String
Private c As Collection

'Untyped set
Public Sub init(arr As Variant)
    isTyped = False
    Set d = New Scripting.Dictionary
    If LBound(arr) = UBound(arr) Then
        add arr(LBound(arr))
    Else
        append arr
    End If
End Sub

' Typed list initialized with a type string
Public Sub initT(typeString As String)
    isTyped = True
    Set d = New Scripting.Dictionary
    innerType = typeString
End Sub
' Lazily typed set which determines the type by looking at the first object added
Public Sub initLT(arr As Variant)
    isTyped = True
    Set d = New Scripting.Dictionary
    If (LBound(arr) = UBound(arr)) Then
        add arr(LBound(arr))
    Else
        append arr
    End If
End Sub

'' ***************************************************************************************
'' Private functions
'' ***************************************************************************************


Private Function createSimilarSetx() As Setx
    Dim s As New Setx
    If isTyped Then
        s.initT innerType
    Else
        s.init Arrays.emptyVariantArray
    End If
    Set createSimilarSetx = l
End Function

Private Function equalsAny(ByVal text As String, ParamArray searchTerms() As Variant) As Boolean
    Dim i As Integer, found As Boolean
    For i = LBound(searchTerms) To UBound(searchTerms)
        found = text = searchTerms(i)
        If found Then Exit For
    Next
    equalsAny = found
End Function

Private Function isTypeCompatible(value As Variant) As Boolean
    If innerType = TypeName(value) Then
        isTypeCompatible = True
    Else
        isTypeCompatible = innerType = "Integer" And equalsAny(TypeName(value), "Byte") _
        Or innerType = "Long" And equalsAny(TypeName(value), "Integer", "Byte") _
        Or innerType = "LongLong" And equalsAny(TypeName(value), "Long", "Integer", "Byte") _
        Or innerType = "Single" And equalsAny(TypeName(value), "LongLong", "Long", "Integer", "Byte") _
        Or innerType = "Double" And equalsAny(TypeName(value), "Single", "LongLong", "Long", "Integer", "Byte") _
        Or innerType = "Currency" And equalsAny(TypeName(value), "Long", "Integer", "Byte", "Single", "Double")
    End If
End Function

Private Sub checkType(value As Variant, member As String)

    If isTyped Then
        If innerType = vbNullString Then
            If TypeName(value) = "Empty" Then
                ' When adding an empty Variant (TypeName() = "Empty") we don't set the typename.
                ' This prevents accidentally trashing the set type by adding an empty variant when
                ' the set was actually meant to hold a real type.
            Else
                innerType = TypeName(value)
            End If
        Else
            If Not isTypeCompatible(value) Then
                Err.Raise E_TYPEMISMATCH, Stringx.format("{0}.{1}", toString, member), _
                      Stringx.format("Type Mismatch. Expected: '{0}' Given: '{1}'", innerType, TypeName(value))
            End If
        End If
    Else
        ' Untyped
    End If

End Sub

Private Function getPrettyInnerTypeName() As String
    If isTyped Then
        If innerType = vbNullString Then
            getPrettyInnerTypeName = "Lazy Unknown"
        Else
            getPrettyInnerTypeName = innerType
        End If
    Else
        getPrettyInnerTypeName = "Untyped"
    End If
End Function


'' ***************************************************************************************
'' Properties
'' ***************************************************************************************

Public Property Get elems() As Long
    elems = d.count
End Property
Public Property Get typed() As Boolean
    typed = isTyped
End Property

Public Property Get innerTypeName() As String
    innerTypeName = innerType
End Property

Public Sub add(value As Variant)
    checkType value, "Setx.add()"
    If Not d.exists(value) Then
        d.add value, 0
    End If
End Sub

Public Sub append(ParamArray values())

        ' Force values to a Variant array.
    Dim varArray() As Variant
    varArray = values
    
    If LBound(varArray) = UBound(varArray) And varType(varArray(LBound(varArray))) = vbObject And TypeName(varArray(LBound(varArray))) = "List" Then
        ' If we have a single List -> unpack it
        
        Dim element As Variant
        For Each element In varArray(LBound(varArray))
            checkType element, "Setx.add()"
            add element
        Next

    ElseIf LBound(varArray) = UBound(varArray) And varType(varArray(LBound(varArray))) = vbObject And TypeName(varArray(LBound(varArray))) = "Collection" Then
        ' If we have a single Collection -> unpack it
        
        Dim element2 As Variant
        For Each element2 In varArray(LBound(varArray))
            checkType element2, "Setx.add()"
            add element2
        Next
        
    ElseIf LBound(varArray) = UBound(varArray) And varType(varArray(LBound(varArray))) = vbObject And TypeName(varArray(LBound(varArray))) = "Setx" Then
        ' If we have a single set -> unpack it
        
        Dim element3 As Variant
        For Each element3 In varArray(LBound(varArray))
            checkType element3, "Setx.add()"
            add element3
        Next
        
    Else
        ' Array case
        If LBound(varArray) = UBound(varArray) And Variants.isArray(varArray(LBound(varArray))) Then
            ' If we have a single array -> unpack it
            
            ' This line kills Excel. Thus write it a little differently.
            ' varArray = varArray(LBound(varArray))
            Dim tmp() As Variant
            tmp = Arrays.toVariantArray(varArray(LBound(varArray)))
            
            varArray = tmp
        End If
        
        Dim i As Long
        For i = LBound(varArray) To UBound(varArray)
            checkType varArray(i), "Setx.add()"
            add varArray(i)
        Next
    End If

End Sub

Public Sub remove(value As Variant)
    On Error Resume Next
    d.remove value
    If Err.number = 32811 Then
        On Error GoTo 0
        Err.Raise E_ARGUMENTOUTOFRANGE, "Setx.remove()", "Setx remove failed on missing element"
    End If
End Sub

Public Sub removeAll()
    d.removeAll
End Sub

Public Function exists(value As Variant) As Boolean
     exists = d.exists(value)

End Function

Public Function equals(other As Setx) As Boolean
    If is_subset(other) Then
        If is_superset(other) Then
            equals = True
        End If
    Else:
        equals = False
    End If
    
    'If d.is_subset(other) And d.is_superset(other) Then
     '   d.equals(other) = True
    'Else:
     '   d.equals(other) = False
    'End If
    
End Function
    
Public Sub discard(value As Variant)
    On Error Resume Next
    d.remove value
    On Error GoTo 0
End Sub

Public Function pick() As Variant

    Dim keys() As Variant
    keys = d.keys()
    
    Dim elementnum As Integer
    Dim random As Double
    random = Rnd(1)
    Dim result As Double
    result = random * elems
    Dim choice As Integer
    choice = Int(result)
    pick = keys(choice)
    
End Function

Public Function pop() As Variant

    If d.count > 0 Then
        Dim keys() As Variant
        keys = d.keys()
        pop = keys(LBound(keys))
        d.remove pop
    End If
    
End Function

Public Function values() As List
    Set values = List_create
    values.append d.keys()

End Function

Public Function union(other As Setx) As Setx
    
    Set union = Setx_create
    Dim element As Variant
    Dim element2 As Variant
    For Each element In other
        union.add element
    
    Next
    
    For Each element2 In Me
        union.add element2
    
    Next

End Function

Public Function intersection(other As Setx) As Setx

    Set intersection = Setx_create
    Dim element As Variant
    For Each element In other
        If d.exists(element) Then
            intersection.add element
        End If
    
    Next
    
End Function

Public Function difference(other As Setx) As Setx

    Set difference = Setx_create
    
    Dim element As Variant
    For Each element In d
        If Not other.exists(element) Then
            difference.add element
        End If
    
    Next
    
End Function

Public Function symmetric_difference(other As Setx) As Setx

    Dim difference1 As Setx: Set difference1 = difference(other)
    Dim difference2 As Setx: Set difference2 = other.difference(Me)
    Set symmetric_difference = difference1.union(difference2)
     
End Function

Public Function is_disjoint(other As Setx) As Boolean

    Dim element As Variant
    is_disjoint = True
    For Each element In other
        If d.exists(element) Then
            is_disjoint = False
            Exit For
        End If
    Next

End Function

Public Function is_subset(other As Setx) As Boolean
    is_subset = True
    Dim element As Variant
    For Each element In Me
        If Not other.exists(element) Then
            is_subset = False
            Exit For
        End If
    Next
    
End Function

Public Function is_superset(other As Setx) As Boolean
    is_superset = True
    Dim element As Variant
    For Each element In other
        If Not d.exists(element) Then
            is_superset = False
            Exit For
        End If
    Next
    
End Function

'Augmented Assigment Methods


Public Function update_union(other As Setx)
    Dim resulting_set As Setx: Set resulting_set = union(other)
    Set d = New Dictionary
    append resulting_set
End Function

Public Function update_intersection(other As Setx)
    Dim resulting_set As Setx: Set resulting_set = intersection(other)
    Set d = New Dictionary
    append resulting_set
    
End Function

Public Function update_difference(other As Setx)
    Dim resulting_set As Setx: Set resulting_set = difference(other)
    Set d = New Dictionary
    append resulting_set
    
End Function

Public Function update_symmetric_difference(other As Setx)
    Dim resulting_set As Setx: Set resulting_set = symmetric_difference(other)
    Set d = New Dictionary
    append resulting_set
    
End Function
Public Property Get newEnum() As IUnknown
Attribute newEnum.VB_UserMemId = -4
    ' If For Each loops return "Method not supported" do the following:
    ' 1. Export this class to a file (right mouse button on the class in the Project Explorer -> Export file...)
    ' 2. Delete the class in the Project Explorer.
    ' 3. In that exported file add the following text above this comment: "Attribute NewEnum.VB_UserMemID = -4"
    ' 4. Import the changed file again (right mouse button in the Project Explorer -> Import file...)
    Set c = New Collection
    Dim e As Variant
    For Each e In d
        c.add e

    Next
    
    Set newEnum = c.[_NewEnum]

End Property

Public Function toString() As String
    toString = Stringx.format("{0}<{1}>", TypeName(Me), getPrettyInnerTypeName)
End Function

Public Function gist() As String
    gist = "[" & Stringx.join(Me.values, ", ") & "]<" & getPrettyInnerTypeName & ">"
End Function


