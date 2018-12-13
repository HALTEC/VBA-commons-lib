Attribute VB_Name = "Test_Collections"
Option Explicit
Option Base 0 ' Default


Private Sub test_all()
    gStart "Collections"
    
    test_deepListToCollection
    test_deepMapToDictionary
    test_deepMapAndListToCollections
    test_deepCollectionToList
    test_deepDictionaryToMap
    test_deepDictionaryAndCollectionToMapAndList
    
    gStop
End Sub


Private Sub test_deepListToCollection()
    gStart "deepListToCollection"
    
    Dim l As List: Set l = List_create(1, 2, 3, 4)
    Dim c As Collection: Set c = Collections.deepListToCollection(l)
    equals c.count, 4, "Element count matches"
    equals c(1), 1, "Elements match"
    equals c(4), 4, "Elements match"
    
    Set l = List_create
    equals Collections.deepListToCollection(l).count, 0, "Empty list"
    
    Set l = List_createLT(1, 2, 3, 4)
    Set c = Collections.deepListToCollection(l)
    equals c.count, 4, "typed List: Element count matches"
    equals c(1), 1, "typed List: Elements match"
    equals c(4), 4, "typed List: Elements match"
    
    gStop
End Sub


Private Sub test_deepMapToDictionary()
    gStart "deepMapToDictionary"
    
    Dim m As Map: Set m = Map_create("a", 1, "b", 2, "c", 3)
    Dim d As Dictionary: Set d = Collections.deepMapToDictionary(m)
    equals d.count, 3, "Element count matches"
    equals d("a"), 1, "Elements match"
    equals d("c"), 3, "Elements match"
    
    Set m = Map_create
    equals Collections.deepMapToDictionary(m).count, 0, "Empty map"
    
    Set m = Map_createLT("a", 1, "b", 2, "c", 3)
    Set d = Collections.deepMapToDictionary(m)
    equals d.count, 3, "typed Map: Element count matches"
    equals d("a"), 1, "typed Map: Elements match"
    equals d("c"), 3, "typed Map: Elements match"
    
    gStop
End Sub


Private Sub test_deepMapAndListToCollections()
    gStart "deepMapAndListToCollections"
    
    Dim l As List: Set l = List_create(1, Map_create("a", 1))
    Dim c As Collection: Set c = Collections.deepListToCollection(l)
    equals c.count, 2, "list element count matches"
    equals TypeName(c(2)), "Dictionary", "list contained types work"
    
    c(2)("a") = 5
    equals l(1)("a"), 1, "list contained Collection types are cloned"
    
    Set l = List_createLT(Map_create("b", 5), Map_create("a", 1))
    Set c = Collections.deepListToCollection(l)
    equals c.count, 2, "typed List: list element count matches"
    equals TypeName(c(2)), "Dictionary", "typed List: list contained types work"
    
    
    Dim m As Map: Set m = Map_create("a", List_create(1, 2))
    Dim d As Dictionary: Set d = Collections.deepMapToDictionary(m)
    equals d.count, 1, "map element count matches"
    equals TypeName(d("a")), "Collection", "map contained types work"
    
    d("a").remove 2
    d("a").add 5
    equals m("a")(1), 2, "map contained Collection types are cloned"
    
    Set m = Map_createLT("a", List_create(1, 2))
    Set d = Collections.deepMapToDictionary(m)
    equals d.count, 1, "typed Map: map element count matches"
    equals TypeName(d("a")), "Collection", "typed Map: map contained types work"
    
    gStop
End Sub


Private Sub test_deepCollectionToList()
    gStart "deepCollectionToList"
    
    Dim c As New Collection
    c.add 1: c.add 2: c.add 3: c.add 4
    Dim l As List: Set l = Collections.deepCollectionToList(c)
    equals l.elems, 4, "Element count matches"
    equals l(0), 1, "Elements match"
    equals l(3), 4, "Elements match"
    
    Set c = New Collection
    equals Collections.deepCollectionToList(c).elems, 0, "Empty list"
    
    gStop
End Sub


Private Sub test_deepDictionaryToMap()
    gStart "deepDictionaryToMap"
    
    Dim d As New Dictionary
    d.add "a", 1: d.add "b", 2: d.add "c", 3
    Dim m As Map: Set m = Collections.deepDictionaryToMap(d)
    equals m.elems, 3, "Element count matches"
    equals m("a"), 1, "Elements match"
    equals m("c"), 3, "Elements match"
    
    Set d = New Dictionary
    equals Collections.deepDictionaryToMap(d).elems, 0, "Empty map"
    
    gStop
End Sub


Private Sub test_deepDictionaryAndCollectionToMapAndList()
    gStart "deepDictionaryAndCollectionToMapAndList"
    
    Dim d As New Dictionary: d.add "a", 1
    Dim c As New Collection: c.add 1: c.add d
    Dim l As List: Set l = Collections.deepCollectionToList(c)
    equals l.elems, 2, "list element count matches"
    equals TypeName(l(1)), "Map", "list contained types work"
    
    l(1)("a") = 5
    equals c(2)("a"), 1, "list contained Collection types are cloned"
    
    
    Set c = New Collection: c.add 1: c.add 2
    Set d = New Dictionary: d.add "a", c
    Dim m As Map: Set m = Collections.deepDictionaryToMap(d)
    equals m.elems, 1, "map element count matches"
    equals TypeName(m("a")), "List", "map contained types work"
    
    m("a").pop
    m("a").push 5
    equals d("a")(2), 2, "map contained Collection types are cloned"
    
    gStop
End Sub

