Attribute VB_Name = "Test_Collections"
Option Explicit
Option Base 0 ' Default


Private Sub test_all()
    gStart "Collections"
    
    test_listToCollection
    test_mapToDictionary
    test_mapAndListToCollections
    
    gStop
End Sub

Private Sub test_listToCollection()
    gStart "listToCollection"
    
    Dim l As List: Set l = List_create(1, 2, 3, 4)
    Dim c As Collection: Set c = Collections.listToCollection(l)
    equals c.count, 4, "Element count matches"
    equals c(1), 1, "Elements match"
    equals c(4), 4, "Elements match"
    
    Set l = List_create
    equals Collections.listToCollection(l).count, 0, "Empty list"
    
    Set l = List_createLT(1, 2, 3, 4)
    Set c = Collections.listToCollection(l)
    equals c.count, 4, "typed List: Element count matches"
    equals c(1), 1, "typed List: Elements match"
    equals c(4), 4, "typed List: Elements match"
    
    gStop
End Sub

Private Sub test_mapToDictionary()
    gStart "mapToDictionary"
    
    Dim m As Map: Set m = Map_create("a", 1, "b", 2, "c", 3)
    Dim d As Dictionary: Set d = Collections.mapToDictionary(m)
    equals d.count, 3, "Element count matches"
    equals d("a"), 1, "Elements match"
    equals d("c"), 3, "Elements match"
    
    Set m = Map_create
    equals Collections.mapToDictionary(m).count, 0, "Empty map"
    
    Set m = Map_createLT("a", 1, "b", 2, "c", 3)
    Set d = Collections.mapToDictionary(m)
    equals d.count, 3, "typed Map: Element count matches"
    equals d("a"), 1, "typed Map: Elements match"
    equals d("c"), 3, "typed Map: Elements match"
    
    gStop
End Sub

Private Sub test_mapAndListToCollections()
    gStart "mapAndListToCollections"
    
    Dim l As List: Set l = List_create(1, Map_create("a", 1))
    Dim c As Collection: Set c = Collections.listToCollection(l)
    equals c.count, 2, "list element count matches"
    equals TypeName(c(2)), "Dictionary", "list contained types work"
    
    c(2)("a") = 5
    equals l(1)("a"), 1, "list contained Collections types are cloned"
    
    Set l = List_createLT(Map_create("b", 5), Map_create("a", 1))
    Set c = Collections.listToCollection(l)
    equals c.count, 2, "typed List: list element count matches"
    equals TypeName(c(2)), "Dictionary", "typed List: list contained types work"
    
    
    Dim m As Map: Set m = Map_create("a", List_create(1, 2))
    Dim d As Dictionary: Set d = Collections.mapToDictionary(m)
    equals d.count, 1, "map element count matches"
    equals TypeName(d("a")), "Collection", "map contained types work"
    
    d("a").remove 2
    d("a").add 5
    equals m("a")(1), 2, "map contained Collections types are cloned"
    
    Set m = Map_createLT("a", List_create(1, 2))
    Set d = Collections.mapToDictionary(m)
    equals d.count, 1, "typed Map: map element count matches"
    equals TypeName(d("a")), "Collection", "typed Map: map contained types work"
    
    gStop
End Sub
