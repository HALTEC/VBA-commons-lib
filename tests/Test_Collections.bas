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
    
    gStop
End Sub

Private Sub test_mapAndListToCollections()
    gStart "mapAndListToCollections"
    
    Dim l As List: Set l = List_create(1, Map_create("a", 1))
    Dim c As Collection: Set c = Collections.listToCollection(l)
    equals c.count, 2, "list element count matches"
    equals TypeName(c(2)), "Dictionary", "list contained types work"
    
    c(2)("a") = 5
    equals l(1)("a"), 1, "list contained Collections types types are cloned"
    
    
    Dim m As Map: Set m = Map_create("a", List_create(1, 2))
    Dim d As Dictionary: Set d = Collections.mapToDictionary(m)
    equals d.count, 2, "map element count matches"
    equals TypeName(d("a")), "Collection", "map contained types work"
    
    ' Fails here:
    d("a")(2) = 5
    equals m("a")(1), 2, "map contained Collections types types are cloned"
    
    
    gStop
End Sub
