VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Collections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function listToCollection(l As List) As Collection
    Dim l2 As List: Set l2 = l.clone
    
    Dim i As Integer
    For i = 0 To l2.elems - 1
        If varType(l2(i)) = vbObject And TypeName(l2(i)) = "Map" Then
            Set l2(i) = mapToDictionary(l2(i))
        ElseIf varType(l2(i)) = vbObject And TypeName(l2(i)) = "List" Then
            Set l2(i) = listToCollection(l2(i))
        End If
    Next
    
    Set listToCollection = l2.innerCollection
End Function

Public Function mapToDictionary(m As Map) As Dictionary
    Dim m2 As Map: Set m2 = m.clone
    
    Dim k As Variant
    For Each k In m2.keys
        If varType(m2(k)) = vbObject And TypeName(m2(k)) = "Map" Then
            Set m2(k) = mapToDictionary(m2(k))
        ElseIf varType(m2(k)) = vbObject And TypeName(m2(k)) = "List" Then
            Set m2(k) = listToCollection(m2(k))
        End If
    Next
    
    Set mapToDictionary = m2.innerDictionary
End Function

