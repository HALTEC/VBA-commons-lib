VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Collections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function deepListToCollection(l As List) As Collection
    Dim l2 As List: Set l2 = List_create: l2.append l
    
    Dim i As Integer
    For i = 0 To l2.elems - 1
        If varType(l2(i)) = vbObject And TypeName(l2(i)) = "Map" Then
            Set l2(i) = deepMapToDictionary(l2(i))
        ElseIf varType(l2(i)) = vbObject And TypeName(l2(i)) = "List" Then
            Set l2(i) = deepListToCollection(l2(i))
        End If
    Next
    
    Set deepListToCollection = l2.innerCollection
End Function

Public Function deepMapToDictionary(m As Map) As Dictionary
    Dim m2 As Map: Set m2 = Map_create(m)
    
    Dim k As Variant
    For Each k In m2.keys
        If varType(m2(k)) = vbObject And TypeName(m2(k)) = "Map" Then
            Set m2(k) = deepMapToDictionary(m2(k))
        ElseIf varType(m2(k)) = vbObject And TypeName(m2(k)) = "List" Then
            Set m2(k) = deepListToCollection(m2(k))
        End If
    Next
    
    Set deepMapToDictionary = m2.innerDictionary
End Function


Public Function deepCollectionToList(c As Collection) As List
    Dim l As List: Set l = List_create: l.append c
    
    Dim i As Integer
    For i = 0 To l.elems - 1
        If varType(l(i)) = vbObject And TypeName(l(i)) = "Dictionary" Then
            Set l(i) = deepDictionaryToMap(l(i))
        ElseIf varType(l(i)) = vbObject And TypeName(l(i)) = "Collection" Then
            Set l(i) = deepCollectionToList(l(i))
        End If
    Next
    
    Set deepCollectionToList = l
End Function

Public Function deepDictionaryToMap(d As Dictionary) As Map
    Dim m As Map: Set m = Map_create(d)
    
    Dim k As Variant
    For Each k In m.keys
        If varType(m(k)) = vbObject And TypeName(m(k)) = "Dictionary" Then
            Set m(k) = deepDictionaryToMap(m(k))
        ElseIf varType(m(k)) = vbObject And TypeName(m(k)) = "Collection" Then
            Set m(k) = deepCollectionToList(m(k))
        End If
    Next
    
    Set deepDictionaryToMap = m
End Function
