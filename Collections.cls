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
    Dim c As Collection: Set c = l.clone.innerCollection
    
    Dim i As Integer
    For i = 1 To c.count
        If varType(c(i)) = vbObject And TypeName(c(i)) = "Dictionary" Then
            Set c(i) = mapToDictionary(c(i))
        ElseIf varType(c(i)) = vbObject And TypeName(c(i)) = "List" Then
            Set c(i) = listToCollection(c(i))
        End If
    Next
    
    Set listToCollection = c
End Function

Public Function mapToDictionary(m As Map) As Dictionary
    Dim d As Dictionary: Set d = m.clone.innerDictionary
    
    Dim k As Variant
    For Each k In d.keys
        If varType(d(k)) = vbObject And TypeName(d(k)) = "Dictionary" Then
            Set d(k) = mapToDictionary(d(k))
        ElseIf varType(d(k)) = vbObject And TypeName(d(k)) = "List" Then
            Set d(k) = listToCollection(d(k))
        End If
    Next
    
    Set mapToDictionary = d
End Function

