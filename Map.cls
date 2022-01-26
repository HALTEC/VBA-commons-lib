VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private d As Scripting.Dictionary
Private isTyped As Boolean
Private keyType As String
Private valType As String

' Untyped map
Public Sub init(arr As Variant)
    isTyped = False
    Set d = New Scripting.Dictionary
    unpackForInsert arr
End Sub

' Typed map initialized with a type string for the key and value.
Public Sub initT(keyTypeString As String, valTypeString As String)
    isTyped = True
    Set d = New Scripting.Dictionary
    
    keyType = keyTypeString
    valType = valTypeString
End Sub

' Lazily typed map which determines the type by looking at the first pair added.
Public Sub initLT(arr As Variant)
    isTyped = True
    Set d = New Scripting.Dictionary
    unpackForInsert arr
End Sub

'*************************'
'******** Private ********'
'*************************'

Private Sub unpackForInsert(arr As Variant)
    If LBound(arr) = UBound(arr) Then
        insert arr(LBound(arr))
    ElseIf LBound(arr) = UBound(arr) - 1 Then
        insert arr(LBound(arr)), arr(UBound(arr))
    Else
        insert arr
    End If
End Sub

Private Function equalsAny(ByVal text As String, ParamArray searchTerms() As Variant) As Boolean
    Dim i As Integer, found As Boolean
    For i = LBound(searchTerms) To UBound(searchTerms)
        found = text = searchTerms(i)
        If found Then Exit For
    Next
    equalsAny = found
End Function

Private Function isTypeCompatible(elem As Variant, elemType As String) As Boolean
    If elemType = TypeName(elem) Then
        isTypeCompatible = True
    Else
        isTypeCompatible = elemType = "Integer" And equalsAny(TypeName(elem), "Byte") _
        Or elemType = "Long" And equalsAny(TypeName(elem), "Integer", "Byte") _
        Or elemType = "LongLong" And equalsAny(TypeName(elem), "Long", "Integer", "Byte") _
        Or elemType = "Single" And equalsAny(TypeName(elem), "LongLong", "Long", "Integer", "Byte") _
        Or elemType = "Double" And equalsAny(TypeName(elem), "Single", "LongLong", "Long", "Integer", "Byte") _
        Or elemType = "Currency" And equalsAny(TypeName(elem), "Long", "Integer", "Byte", "Single", "Double")
    End If
End Function

Private Function bothTypeCompatible(key As Variant, val As Variant) As Boolean
    bothTypeCompatible = isTypeCompatible(key, keyType) And isTypeCompatible(val, valType)
End Function

Private Sub checkType(key As Variant, val As Variant, member As String)
    If isTyped Then
        If keyType = vbNullString Then
            If TypeName(key) <> "Empty" And TypeName(val) <> "Empty" Then
                keyType = TypeName(key)
                valType = TypeName(val)
            End If
        ElseIf Not bothTypeCompatible(key, val) Then
            Err.Raise E_TYPEMISMATCH, Stringx.format("{0}.{1}", toString, member), _
                Stringx.format("Type Mismatch. Expected <{0}, {1}> Given: <{2}, {3}>", _
                keyType, valType, TypeName(key), TypeName(val))
        End If
    End If
End Sub

Private Function getPrettyTypeName() As String
    If isTyped Then
        If keyType = vbNullString Then
            getPrettyTypeName = "<Lazy Unknown>"
        Else
            getPrettyTypeName = "<" & keyType & ", " & valType & ">"
        End If
    Else
        getPrettyTypeName = "<Untyped>"
    End If
End Function

Private Function getList(typeString As String) As List
    Dim l As List
    If isTyped Then
        Set l = List_createT(typeString)
    Else
        Set l = List_create
    End If
    
    Set getList = l
End Function

Private Sub Class_Terminate()
    Set d = Nothing
End Sub

'************************'
'****** Properties ******'
'************************'

Public Property Get item(ByVal key As Variant) As Variant
    If IsObject(d(key)) Then
        Set item = d(key)
    Else
        item = d(key)
    End If
End Property
Public Property Set item(ByVal key As Variant, ByVal value As Variant)
    checkType key, value, "Set()"
    Set d(key) = value
End Property
Public Property Let item(ByVal key As Variant, ByVal value As Variant)
Attribute item.VB_UserMemId = 0
    checkType key, value, "Let()"
    d(key) = value
End Property

Public Property Get elems() As Long
    elems = d.count
End Property

Public Property Get typed() As Boolean
    typed = isTyped
End Property

Public Property Get keyTypeName() As String
    keyTypeName = keyType
End Property

Public Property Get valTypeName() As String
    valTypeName = valType
End Property

'************************'
'******** Public ********'
'************************'

Public Function exists(key As Variant) As Boolean
    exists = d.exists(key)
End Function

Public Sub remove(key As Variant)
    d.remove key
End Sub

Public Sub removeAll()
    d.removeAll
End Sub

Public Sub add(key As Variant, val As Variant)
    checkType key, val, "Add()"
    d.add key, val
End Sub

Public Sub insert(ParamArray args())
    Dim varArray() As Variant: varArray = args
    Dim i As Long
    If LBound(varArray) = UBound(varArray) And varType(varArray(LBound(varArray))) = vbObject And TypeName(varArray(LBound(varArray))) = "List" Then
        Dim l As List: Set l = varArray(LBound(varArray))
        If l.elems Mod 2 <> 0 Then
            Err.Raise E_INVALIDINPUT, "Map.Insert()", "Incomplete key/value pair"
        End If
        For i = 0 To l.elems - 1
            If i Mod 2 = 0 And i + 1 < l.elems Then
                checkType l(i), l(i + 1), "Insert()"
                If IsObject(l(i + 1)) Then
                    Set d(l(i)) = l(i + 1)
                Else
                    d(l(i)) = l(i + 1)
                End If
            End If
        Next
    ElseIf LBound(varArray) = UBound(varArray) And varType(varArray(LBound(varArray))) = vbObject And TypeName(varArray(LBound(varArray))) = "Dictionary" Then
        Dim key As Variant
        For Each key In varArray(LBound(varArray)).keys
            checkType key, varArray(LBound(varArray))(key), "insert()"
            If IsObject(varArray(LBound(varArray))(key)) Then
                ' The .item is necessary, because VBA throws a "argument is not optional" error otherwise. Not entirely sure
                ' where VBA sees a function call that's missing parameters though...
                Set d(key) = varArray(LBound(varArray)).item(key)
            Else
                d(key) = varArray(LBound(varArray)).item(key)
            End If
        Next key
    ElseIf LBound(varArray) = UBound(varArray) And varType(varArray(LBound(varArray))) = vbObject And TypeName(varArray(LBound(varArray))) = "Map" Then
        Dim kv As KeyValuePair
        For Each kv In varArray(LBound(varArray)).pairs
            checkType kv.key, kv.value, "insert()"
            If IsObject(kv.value) Then
                Set d(kv.key) = kv.value
            Else
                d(kv.key) = kv.value
            End If
        Next
    ElseIf LBound(varArray) = UBound(varArray) - 1 And varType(varArray(LBound(varArray))) = vbObject _
            And TypeName(varArray(LBound(varArray))) = "List" And varType(varArray(UBound(varArray))) = vbObject _
            And TypeName(varArray(UBound(varArray))) = "List" Then
        Dim keys As List: Set keys = varArray(LBound(varArray))
        Dim values As List: Set values = varArray(UBound(varArray))
        If keys.elems <> values.elems Then
            Err.Raise E_INVALIDINPUT, "Map.Insert()", "Different number of keys and values"
        End If
        For i = 0 To keys.elems - 1
            checkType keys(i), values(i), "Insert()"
            If IsObject(values(i)) Then
                Set d(keys(i)) = values(i)
            Else
                d(keys(i)) = values(i)
            End If
        Next
    ElseIf LBound(varArray) = UBound(varArray) - 1 And Variants.isArray(varArray(LBound(varArray))) _
            And Variants.isArray(varArray(UBound(varArray))) Then
        Dim keysArr As Variant: keysArr = varArray(LBound(varArray))
        Dim valuesArr As Variant: valuesArr = varArray(UBound(varArray))
        If UBound(keysArr) - LBound(keysArr) <> UBound(valuesArr) - LBound(valuesArr) Then
            Err.Raise E_INVALIDINPUT, "Map.Insert()", "Different number of keys and values"
        End If
        For i = LBound(keysArr) To UBound(keysArr)
            checkType keysArr(i), valuesArr(i - LBound(keysArr) + LBound(valuesArr)), "Insert()"
            If IsObject(valuesArr(i - LBound(keysArr) + LBound(valuesArr))) Then
                Set d(keysArr(i)) = valuesArr(i - LBound(keysArr) + LBound(valuesArr))
            Else
                d(keysArr(i)) = valuesArr(i - LBound(keysArr) + LBound(valuesArr))
            End If
        Next
    Else
        If LBound(varArray) = UBound(varArray) And Variants.isArray(varArray(LBound(varArray))) Then
            Dim tmp() As Variant
            tmp = Arrays.toVariantArray(varArray(LBound(varArray)))
            
            varArray = tmp
        End If
        If (LBound(varArray) - UBound(varArray)) Mod 2 = 0 Then
            Err.Raise E_INVALIDINPUT, "Map.Insert()", "Incomplete key/value pair"
        End If
        For i = LBound(varArray) To UBound(varArray)
            If i Mod 2 = 0 And i + 1 <= UBound(varArray) Then
                checkType varArray(i), varArray(i + 1), "Insert()"
                If IsObject(varArray(i + 1)) Then
                    Set d(varArray(i)) = varArray(i + 1)
                Else
                    d(varArray(i)) = varArray(i + 1)
                End If
            End If
        Next
    End If
End Sub

Public Function keys() As List
    Dim l As List: Set l = getList(keyType)
    l.append d.keys
    Set keys = l
End Function

Public Function values() As List
    Dim l As List: Set l = getList(valType)
    Dim key As Variant
    For Each key In d.keys
        l.push d(key)
    Next key
    Set values = l
End Function

Public Function pairs() As List
    Dim l As List: Set l = List_create
    Dim key As Variant
    Dim pair As KeyValuePair
    For Each key In d.keys
        Set pair = New KeyValuePair
        If IsObject(key) Then
            Set pair.key = key
        Else
            pair.key = key
        End If
        If IsObject(d(key)) Then
            Set pair.value = d(key)
        Else
            pair.value = d(key)
        End If
        l.push pair
    Next key
    Set pairs = l
End Function
Public Function antiPairs() As List
    Dim l As List: Set l = List_create
    Dim key As Variant
    Dim pair As KeyValuePair
    For Each key In d.keys
        Set pair = New KeyValuePair
        If IsObject(d(key)) Then
            Set pair.key = d(key)
        Else
            pair.key = d(key)
        End If
        If IsObject(key) Then
            Set pair.value = key
        Else
            pair.value = key
        End If
        l.push pair
    Next key
    Set antiPairs = l
End Function
Public Function sort(Optional sortOrder As sortOrder = ascending) As List
    Dim keyArray As Variant: keyArray = d.keys
    On Error Resume Next
    Arrays.sort keyArray, sortOrder
    If Err.number = E_ARGUMENTOUTOFRANGE Then
        On Error GoTo 0
        Err.Raise E_ILLEGALSTATE, "Map.sort()", "Map keys are not sortable."
    ElseIf Err.number = E_TYPEMISMATCH Then
        On Error GoTo 0
        Err.Raise E_ILLEGALSTATE, "Map.sort()", "Map keys are of different type and are thus uncomparable."
    ElseIf Err.number <> 0 Then
        On Error GoTo 0
        Err.Raise Err.number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
    Dim pairList As List: Set pairList = List_create
    Dim pair As KeyValuePair
    Dim key As Variant
    For Each key In keyArray
        Set pair = New KeyValuePair
        If IsObject(key) Then
            Set pair.key = key
        Else
            pair.key = key
        End If
        If IsObject(d(key)) Then
            Set pair.value = d(key)
        Else
            pair.value = d(key)
        End If
        pairList.push pair
    Next key
    Set sort = pairList
End Function

Public Function clone() As Map
    Dim m As Map
    If isTyped Then
        Set m = Map_createT(keyType, valType)
    Else
        Set m = Map_create()
    End If
    m.insert d
    Set clone = m
End Function

Public Function innerDictionary() As Dictionary
    Set innerDictionary = d
End Function

Public Function toString() As String
    toString = TypeName(Me) & getPrettyTypeName
End Function
Public Function gist() As String
    Dim result As String
    Dim keyList As List: Set keyList = keys
    Dim sorted As List: Set sorted = keyList.sort()
    Dim key As Variant
    Dim value As Variant
    result = "{"
    Dim i As Long: i = 0
    For Each key In sorted
        If i > 0 Then
            result = result & ", "
        End If
        
        'If we don't have a type that nicely stringifies, just use the gist.
        If IsObject(d(key)) Then
            value = Variants.gist(d(key))
        ElseIf Not IsNumeric(d(key)) And varType(d(key)) <> vbString Then
            value = Variants.gist(d(key))
        Else
            value = d(key)
        End If
        
        result = result & Stringx.format("{0} => {1}", key, value)
        i = i + 1
        If i >= 100 Then
            result = result & ", (...)"
            Exit For
        End If
    Next key
    result = result & "}"
    gist = result
End Function

Public Function equals(other As Map) As Boolean
    If d.count <> other.elems Then
        equals = False
        Exit Function
    End If
    
    If keyType <> other.keyTypeName Or valType <> other.valTypeName Then
        equals = False
        Exit Function
    End If
    
    Dim key As Variant
    For Each key In d.keys
        If Not Variants.equals(d(key), other(key)) Then
            equals = False
            Exit Function
        End If
    Next key
    
    equals = True
End Function
