VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 0 ' Default

Private c As Collection
Private isTyped As Boolean
Private innerType As String



' Untyped list
Public Sub init(arr As Variant)
    isTyped = False
    Set c = New Collection

    If LBound(arr) = UBound(arr) Then
        push arr(LBound(arr))
    Else
        append arr
    End If
End Sub

' Typed list initialized with a type string.
Public Sub initT(typeString As String)
    isTyped = True
    Set c = New Collection
    
    innerType = typeString
End Sub

' Lazily typed list which determines the type by looking at the first object added.
Public Sub initLT(arr As Variant)
    isTyped = True
    Set c = New Collection
    
    If (LBound(arr) = UBound(arr)) Then
        push arr(LBound(arr))
    Else
        append arr
    End If
End Sub

'' ***************************************************************************************
'' Private functions
'' ***************************************************************************************


Private Function createSimilarList() As List
    Dim l As New List
    If isTyped Then
        l.initT innerType
    Else
        l.init Arrays.emptyVariantArray
    End If
    Set createSimilarList = l
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
'
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
                ' This prevents accidentally trashing the list type by addin an empty variant when
                ' the list was actually meant to hold a real type.
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

Private Sub addBefore(ByVal index As Long, ByRef value As Variant)

    If index < 0 Then
        index = c.count + index
    End If

    If c.count = 0 Then
        c.add value
    ElseIf index = 0 Then
        c.add value, before:=1
    Else
        c.add value, after:=index
    End If
    
End Sub

Private Sub checkIndex(reporter As String, ByVal index As Long, Optional ByVal thisElems As Long = 0, Optional ByVal afterLast As Boolean = False)
    If thisElems < 0 Then
        Err.Raise E_INDEXOUTOFRANGE, reporter, "elems must not be smaller than 0."
    End If
    
    If index < 0 Or _
            afterLast And index > c.count Or _
            Not afterLast And index >= c.count Then
        Err.Raise E_INDEXOUTOFRANGE, reporter, "index (" & index & ") must be in range 0 - " & c.count & "."
    End If
    
    If afterLast And index + thisElems > c.count + 1 Or _
            Not afterLast And index + thisElems > c.count Then
        Err.Raise E_INDEXOUTOFRANGE, reporter, "index + elems must not be > List.elems."
    End If
End Sub

Private Sub Class_Terminate()
    Set c = Nothing
End Sub

'' ***************************************************************************************
'' Properties
'' ***************************************************************************************

Public Property Get item(ByVal index As Long) As Variant

    ' Convert "count from the back" negative index to positive.
    If index < 0 Then
        index = c.count + index
    End If
    
    checkIndex "List()", index
    
    ' Make index 1-based
    index = index + 1
    

    If IsObject(c(index)) Then
        Set item = c(index)
    Else
        item = c(index)
    End If

End Property

Public Property Let item(ByVal index As Long, ByVal value As Variant)
Attribute item.VB_UserMemId = 0
    ' If item access ( myList(3) ) returns "Method not supported" do the following:
    ' 1. Export this class to a file (right mouse button on the class in the Project Explorer -> Export file...)
    ' 2. Delete the class in the Project Explorer.
    ' 3. In that exported file add the following text above this comment: "Attribute NewEnum.VB_UserMemID = 0"
    ' 4. Import the changed file again (right mous button in the Project Explorer -> Import file...)

    checkType value, "List()="
    
    ' Convert "count from the back" negative index to positive.
    If index < 0 Then
        index = c.count + index
    End If
    
    checkIndex "List()=", index

    c.remove index + 1
    
    addBefore index, value

End Property

Public Property Set item(ByVal index As Long, ByVal value As Variant)

    checkType value, "List()="
    
    ' Convert "count from the back" negative index to positive.
    If index < 0 Then
        index = c.count + index
    End If
    
    checkIndex "List()=", index

    c.remove index + 1
    
    addBefore index, value

End Property

Public Property Get newEnum() As IUnknown
Attribute newEnum.VB_UserMemId = -4
    ' If For Each loops return "Method not supported" do the following:
    ' 1. Export this class to a file (right mouse button on the class in the Project Explorer -> Export file...)
    ' 2. Delete the class in the Project Explorer.
    ' 3. In that exported file add the following text above this comment: "Attribute NewEnum.VB_UserMemID = -4"
    ' 4. Import the changed file again (right mous button in the Project Explorer -> Import file...)
    Set newEnum = c.[_NewEnum]

End Property

Public Property Get elems() As Long

    elems = c.count

End Property

Public Property Get typed() As Boolean
    typed = isTyped
End Property

Public Property Get innerTypeName() As String
    innerTypeName = innerType
End Property

'' ***************************************************************************************
'' Public functions
'' ***************************************************************************************

Public Function pop() As Variant

    If c.count = 0 Then
        Err.Raise E_ILLEGALSTATE, "List.pop()", "Can't .pop an empty list"
    End If
    
    If IsObject(c(c.count)) Then
        Set pop = c(c.count)
    Else
        pop = c(c.count)
    End If

    c.remove c.count
    
End Function

Public Sub push(value As Variant)

    checkType value, "List.push()"
    c.add value
    
End Sub

Public Sub append(ParamArray values())

    ' Force values to a Variant array.
    Dim varArray() As Variant
    varArray = values
    
    If LBound(varArray) = UBound(varArray) And varType(varArray(LBound(varArray))) = vbObject And TypeName(varArray(LBound(varArray))) = "List" Then
        ' If we have a single List -> unpack it
        
        Dim element As Variant
        For Each element In varArray(LBound(varArray))
            checkType element, "List.append()"
            c.add element
        Next
        
    ElseIf LBound(varArray) = UBound(varArray) And varType(varArray(LBound(varArray))) = vbObject And TypeName(varArray(LBound(varArray))) = "Collection" Then
        ' If we have a single Collection -> unpack it
        
        Dim element2 As Variant
        For Each element2 In varArray(LBound(varArray))
            checkType element2, "List.append()"
            c.add element2
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
            checkType varArray(i), "List.append()"
            c.add varArray(i)
        Next
    End If

End Sub

Public Function shift() As Variant

    If c.count = 0 Then
        Err.Raise E_ILLEGALSTATE, "List.shift()", "Can't .shift an empty list"
    End If
    
    If IsObject(c(1)) Then
        Set shift = c(1)
    Else
        shift = c(1)
    End If
    
    c.remove 1
End Function

Public Sub unshift(value As Variant)

    checkType value, "List.unshift()"
    
    If c.count = 0 Then
        c.add value
    Else
        c.add value, before:=1
    End If
    
End Sub

Public Sub prepend(ParamArray values())

    Dim pos As Long

    ' Force values to a Variant array.
    Dim varArray() As Variant
    varArray = values
    
    If LBound(varArray) = UBound(varArray) And varType(varArray(LBound(varArray))) = vbObject And TypeName(varArray(LBound(varArray))) = "List" Then
        ' If we have a single list -> unpack it
        
        pos = 0
        Dim element As Variant
        For Each element In varArray(LBound(varArray))
            checkType element, "List.prepend()"
            
            addBefore pos, element
            
            pos = pos + 1
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
        pos = 0
        For i = LBound(varArray) To UBound(varArray)
            checkType varArray(i), "List.prepend()"
            addBefore pos, varArray(i)
            pos = pos + 1
        Next
    End If
End Sub

Public Function splice(ByVal index As Long, ByVal elems As Long, Optional ByVal replacement As Variant) As List

    Dim before As List

    If index < 0 Then
        index = c.count + index
    End If
    
    checkIndex "List.splice()", index, elems, True
    
    If elems > 0 Then
        Set splice = getRange(index, elems)
        remove index, elems
    Else
        Set splice = createSimilarList
    End If
    
    If IsMissing(replacement) Then
        ' Done. Nothing to do.
    ElseIf Variants.isArray(replacement) Then
        ' Force values to a Variant array.
        Dim varArray() As Variant
        varArray = replacement
        
        Dim i As Long
        For i = LBound(varArray) To UBound(varArray)
        
            addBefore index, varArray(i)
            index = index + 1
            
        Next
    ElseIf varType(replacement) = vbObject And TypeName(replacement) = "List" Then
        Dim element As Variant
        For Each element In replacement
        
            addBefore index, element
            index = index + 1
            
        Next
    Else
        Err.Raise E_TYPEMISMATCH, "List.splice()", "replacement must be either an array or a List."
    End If
    
End Function

Public Function clone() As List
    Set clone = getRange(0, c.count)
End Function

Public Function getRange(ByVal index As Long, ByVal elems As Long) As List
    Dim result As List: Set result = createSimilarList
    
    If elems > 0 Then
    
        If index < 0 Then
            index = c.count + index
        End If
        
        checkIndex "List.getRange()", index, elems

        
        ' 1-based inclusive lastIndex
        Dim lastIndex As Long
        lastIndex = index + elems
        
        If lastIndex > c.count Then
            Err.Raise E_INDEXOUTOFRANGE, "List.getRange()", "End index > size"
        End If
        
        Dim i As Long
        For i = index + 1 To lastIndex
            result.push c(i)
        Next
    End If
    
    Set getRange = result

End Function

Public Sub remove(ByVal index As Long, Optional ByVal elems As Long = 1)
    
    If index < 0 Then
        index = c.count + index
    End If
    
    checkIndex "List.remove()", index, elems

    Dim i As Long
    For i = index + elems To index + 1 Step -1
        c.remove i
    Next

End Sub

Public Function first() As Variant
    
    If c.count = 0 Then
        Err.Raise E_ILLEGALSTATE, "List.first()", "Can not call first on an empty list"
    End If

    If IsObject(c(1)) Then
        Set first = c(1)
    Else
        first = c(1)
    End If

End Function

Public Function last() As Variant

    If c.count = 0 Then
        Err.Raise E_ILLEGALSTATE, "List.last()", "Can not call last on an empty list"
    End If

    If IsObject(c(c.count)) Then
        Set last = c(c.count)
    Else
        last = c(c.count)
    End If

End Function

Public Sub clear()

    Do Until c.count = 0
        c.remove 1
    Loop

End Sub

Public Function contains(value As Variant) As Boolean

    contains = indexOf(value) <> -1

End Function

Public Function indexOf(value As Variant) As Long

    If c.count = 0 Then indexOf = -1: Exit Function
    
    Dim i As Long
    For i = 1 To c.count
        If Variants.equals(value, c(i)) Then
            indexOf = i - 1
            Exit Function
        End If
    Next

    indexOf = -1

End Function

Public Function lastIndexOf(value As Variant) As Long

    If c.count = 0 Then lastIndexOf = -1: Exit Function

    Dim i As Long
    For i = c.count To 1 Step -1
        If Variants.equals(value, c(i)) Then
            lastIndexOf = i - 1
            Exit Function
        End If
    Next
    
    lastIndexOf = -1

End Function

'' ***************************************************************************************
'' Auxiliary functions
'' ***************************************************************************************

Public Function max() As Variant

    Dim largest As Variant
    Dim isLarger As Boolean

    Dim i As Long
    For i = 0 To c.count - 1

        If IsObject(item(i)) Then

            If IsEmpty(largest) Then Set largest = item(i)
            isLarger = Math.cmp(item(i), largest) > 0

            If isLarger Or IsEmpty(max) Then
                Set largest = item(i)
                Set max = largest
            End If

        Else

            If IsEmpty(largest) Then largest = item(i)
            isLarger = Math.cmp(item(i), largest) > 0

            If isLarger Or IsEmpty(max) Then
                largest = item(i)
                max = largest
            End If

        End If

    Next

End Function

Public Function min() As Variant

    Dim smallest As Variant
    Dim isSmaller As Boolean

    Dim i As Long
    For i = 0 To c.count - 1

        If IsObject(item(i)) Then

            If IsEmpty(smallest) Then Set smallest = item(i)
            isSmaller = Math.cmp(item(i), smallest) < 0

            If isSmaller Or IsEmpty(min) Then
                Set smallest = item(i)
                Set min = smallest
            End If

        Else

            If IsEmpty(smallest) Then smallest = item(i)
            isSmaller = Math.cmp(item(i), smallest) < 0

            If isSmaller Or IsEmpty(min) Then
                smallest = item(i)
                min = smallest
            End If

        End If

    Next

End Function

Public Function reverse()
Attribute reverse.VB_Description = "Reverses the order of the elements in the entire List."

    Dim newList As List: Set newList = createSimilarList
    
    Do Until c.count = 0
        newList.push pop
    Loop
    
    Set reverse = newList

End Function

Public Function sort(Optional sortOrder As sortOrder = ascending) As List
Attribute sort.VB_Description = "Sorts the elements in the entire List."

    Dim newList As List: Set newList = createSimilarList

    If c.count > 0 Then
    
        Dim a As Variant
        a = toArray
        
        On Error Resume Next
        Arrays.sort a, sortOrder
        If Err.number = E_ARGUMENTOUTOFRANGE Then
            On Error GoTo 0
            Err.Raise E_ILLEGALSTATE, "List.sort()", "List elements are not sortable."
        ElseIf Err.number = E_TYPEMISMATCH Then
            On Error GoTo 0
            Err.Raise E_ILLEGALSTATE, "List.sort()", "List elements are of different type and are thus uncomparable."
        ElseIf Err.number <> 0 Then
            On Error GoTo 0
            Err.Raise Err.number, Err.source, Err.description, Err.HelpFile, Err.HelpContext
        End If
        
        newList.append a
    
    End If
    
    Set sort = newList

End Function


Public Function uniq() As List

    Dim newList As List: Set newList = createSimilarList

    If c.count > 0 Then
        Dim elem As Variant
        For Each elem In Me
            If Not newList.contains(elem) Then
                newList.push elem
            End If
        Next
    End If
    
    Set uniq = newList
End Function


Public Function equals(other As List) As Boolean

    If c.count <> other.elems Then
        equals = False
        Exit Function
    End If

    If innerType <> other.innerTypeName Then
        equals = False
        Exit Function
    End If

    Dim i As Long
    For i = 0 To c.count - 1

        If Not Variants.equals(item(i), other(i)) Then
            equals = False
            Exit Function
        End If

    Next
    
    equals = True

End Function

Public Function toArray() As Variant()
Attribute toArray.VB_Description = "Copies the elements of the List to a new array."

    If c.count = 0 Then Exit Function

    Dim result() As Variant
    ReDim result(0 To c.count - 1)

    Dim i As Long
    For i = 0 To c.count - 1
        If IsObject(item(i)) Then
            Set result(i) = item(i)
        Else
            result(i) = item(i)
        End If
    Next

    toArray = result

End Function

Public Function innerCollection() As Collection
    Set innerCollection = c
End Function

Public Function toString() As String
Attribute toString.VB_Description = "Returns a string that represents the current List object."
    toString = Stringx.format("{0}<{1}>", TypeName(Me), getPrettyInnerTypeName)
End Function

Public Function gist() As String
    gist = "[" & Stringx.join(Me, ", ") & "]<" & getPrettyInnerTypeName & ">"
End Function
