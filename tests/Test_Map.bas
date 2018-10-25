Attribute VB_Name = "Test_Map"
Option Explicit

Private Sub test_all()
    gStart "Map"
    
    test_create
    test_createT
    test_createLT
    test_add
    test_keys
    test_values
    test_pairs
    test_antipairs
    test_tostring
    test_gist
    test_sort
    test_exists
    test_remove
    test_removeall
    test_insert
    test_clone
    test_keyvaltypename
    test_equals
    
    gStop
End Sub

Private Sub test_create()
    gStart "create"

    On Error Resume Next
    Map_create
    checkNoError "Creating an empty untyped map"
    On Error GoTo 0
    
    Dim m As Map: Set m = Map_create
    equals m.elems, 0, "Empty untyped map is empty"
    
    Dim m2 As Map: Set m2 = Map_create("a", 1, "b", 2)
    equals m2.elems, 2, "Untyped map can be initialized with values"
    
    Dim m3 As Map: Set m3 = Map_create(m, "bob", "bla", m2)
    equals m3(m), "bob", "Untyped map works with object keys"
    
    Dim d As New Scripting.Dictionary
    d("a") = 1
    d("b") = 2
    Dim m4 As Map: Set m4 = Map_create(d)
    equals m4.elems, 2, "Untyped map can be initialized with a Dictionary"
    
    Dim l As List: Set l = List_create("a", "b", "c")
    Dim l2 As List: Set l2 = List_create(1, 2, 3)
    Dim m5 As Map: Set m5 = Map_create(l, l2)
    equals m5("a"), 1, "Untyped map can be initialized with lists for keys and values"
    
    gStop
End Sub

Private Sub test_createT()
    gStart "createT"
    
    On Error Resume Next
    Map_createT "Integer", "String"
    checkNoError "Creating an empty typed Map"
    On Error GoTo 0
    
    Dim m As Map: Set m = Map_createT("Integer", "String")
    equals m.elems, 0, "Empty typed Map is empty"
    
    Dim m2 As Map: Set m2 = Map_createT("Integer", "String")
    m2.add 2, "asdf"
    m2.add 3, "sdf"
    equals m2.elems, 2, "Filling typed list works"
    equals m2.item(3), "sdf", "Retrieving value from typed list works"
    
    Dim m3 As Map: Set m3 = Map_createT("Integer", "String")
    On Error Resume Next
    m3.add "sdfsf", 5
    checkError E_TYPEMISMATCH, "Typed Map fails on inserting wrong type"
    On Error GoTo 0
    
    gStop
End Sub

Private Sub test_createLT()
    gStart "createLT"
    
    On Error Resume Next
    Map_createLT
    checkNoError "Creating an empty lazy typed Map"
    On Error GoTo 0
    
    Dim m As Map: Set m = Map_createLT
    equals m.elems, 0, "Empty lazy typed Map is empty"
    
    Dim m2 As Map: Set m2 = Map_createLT
    m2.add 2, "asdf"
    m2.add 3, "sdf"
    equals m2.elems, 2, "Filling lazy typed list works"
    equals m2(3), "sdf", "Retrieving value from lazy typed list works"
    m2(2) = "sdfsdf"
    m2(5) = "sdfsfsdf"
    equals m2.elems, 3, "Replacing elements in lazy typed list works"
    equals m2(5), "sdfsfsdf", "Filling lazy typed list with default method works"

    Dim m3 As Map: Set m3 = Map_createLT()
    On Error Resume Next
    m3.add 6, "sfdf"
    m3.add "sdfsf", 5
    checkError E_TYPEMISMATCH, "Lazy typed Map fails on inserting wrong type"
    On Error GoTo 0
    
    Dim m4 As Map: Set m4 = Map_createLT("a", 1, "b", 2)
    equals m4.elems, 2, "Lazy typed Map can be initialized with values"
    
    gStop
End Sub

Private Sub test_add()
    gStart "add"
    
    Dim m As Map: Set m = Map_create
    m.add 20, "hi"
    equals m.elems, 1, "Map contains one element"
    equals m(20), "hi", "Value can be retrieved from Map after Add()ing it"
    
    gStop
End Sub

Private Sub test_keys()
    gStart "keys"
    
    Dim m As Map: Set m = Map_create
    m("a") = 1
    m(1) = "121"
    m("ffd") = 312
    equals m.keys.elems, 3, "Keys() retrieves all keys (1)"
    equals m.keys.contains("ffd"), True, "Keys() retrieves all keys (2)"
    
    gStop
End Sub

Private Sub test_values()
    gStart "values"
    
    Dim m As Map: Set m = Map_create
    m("a") = 1
    m(1) = "121"
    m("ffd") = 312
    equals m.values.elems, 3, "Values() retrieves all values (1)"
    equals m.values.contains("121"), True, "Values() retrieves all values (2)"
    
    gStop
End Sub

Private Sub test_pairs()
    gStart "pairs"
    
    Dim m As Map: Set m = Map_create
    m("a") = 1
    m(42) = "sdfsf"
    Set m("b") = New KeyValuePair
    Dim l As List: Set l = m.pairs()
    equals l.elems, 3, "Pairs() returns all key/value pairs in Map (1)"
    equals l(0).key = "a" Or l(0).key = 42, True, "Pairs() returns all key/value pairs in Map (2)"
    
    gStop
End Sub

Private Sub test_antipairs()
    gStart "antipairs"
    
    Dim m As Map: Set m = Map_create
    m("a") = 1
    m(42) = "sdfsf"
    Dim l As List: Set l = m.antiPairs()
    equals l.elems, 2, "antiPairs() returns all value/key pairs in Map (1)"
    equals l(0).key = 1 Or l(0).key = "sdfsf", True, "antiPairs() returns all value/key pairs in Map (2)"
    
    gStop
End Sub

Private Sub test_tostring()
    gStart "toString"
    
    Dim m As Map: Set m = Map_create
    equals m.toString, "Map<Untyped>", "Untyped toString() works"
    Dim m1 As Map: Set m1 = Map_createT("Integer", "String")
    equals m1.toString, "Map<Integer, String>", "Typed toString() works"
    Dim m2 As Map: Set m2 = Map_createLT
    equals m2.toString, "Map<Lazy Unknown>", "Lazy typed toString() without elements works"
    m2("a") = 1
    equals m2.toString, "Map<String, Integer>", "Lazy typed toString() with elements works"
    
    gStop
End Sub

Private Sub test_gist()
    gStart "gist"
    
    Dim m As Map: Set m = Map_create
    m("a") = 1
    equals m.gist, "a => 1", "gist() works as expected"
    
    gStop
End Sub

Private Sub test_sort()
    gStart "sort"
    
    Dim m As Map: Set m = Map_create
    m("a") = 1
    m(42) = "sdfsf"
    m("sdfs") = 54
    Dim l As List: Set l = m.sort()
    equals l(0).key, 42, "Sort() works as expected (1)"
    equals l(1).value, 1, "Sort() works as expected (2)"
    equals l(2).key, "sdfs", "Sort() works as expected (3)"
    
    gStop
End Sub

Private Sub test_exists()
    gStart "exists"
    
    Dim m As Map: Set m = Map_create
    m("dfsdf") = 324
    m(4) = "sdfs"
    equals m.exists("dfsdf"), True, "Exists() returns True if key exists"
    equals m.exists(5), False, "Exists() return False if key doesn't exist"
    
    gStop
End Sub

Private Sub test_remove()
    gStart "remove"
    
    Dim m As Map: Set m = Map_create
    m("a") = 1
    m("b") = 2
    equals m.elems, 2, "Remove() works as expected (1)"
    m.remove "a"
    equals m.elems, 1, "Remove() works as expected (2)"
    equals m.exists("a"), False, "Remove() works as expected (3)"
    
    gStop
End Sub

Private Sub test_removeall()
    gStart "removeall"
    
    Dim m As Map: Set m = Map_create
    m("a") = 1
    m("b") = 2
    equals m.elems, 2, "RemoveAll() works as expected (1)"
    m.removeAll
    equals m.elems, 0, "RemoveAll() works as expected (2)"
    
    gStop
End Sub

Private Sub test_insert()
    gStart "insert"
    
    Dim m As Map: Set m = Map_create
    m.insert "a", 1, "b", 2
    equals m("a"), 1, "Insert() works with keys and values given directly as arguments"
    Dim l As List: Set l = List_create("c", 3, "d", 4)
    m.insert l
    equals m("d"), 4, "Insert() works with keys and values given as single List"
    Dim k As List: Set k = List_create("e", "f")
    Dim v As List: Set v = List_create(5, 6)
    m.insert k, v
    equals m("e"), 5, "Insert() works with keys and values given as separate Lists"
    Dim a As Variant: a = Array("g", 7, "h", 8)
    m.insert a
    equals m("h"), 8, "Insert() works with keys and values given as single array"
    Dim kArr As Variant: kArr = Array("i", "j")
    Dim vArr As Variant: vArr = Array(9, 10)
    m.insert kArr, vArr
    equals m("i"), 9, "Insert() works with keys and values given as separate arrays"
    
    Dim d As New Scripting.Dictionary
    d("a") = 5
    d(0) = "sdfd"
    m.insert d
    equals m("a"), 5, "insert() correctly adds key/value pairs from dictionary"
    
    On Error Resume Next
    Dim m1 As Map: Set m1 = Map_create
    m1.insert "a", 1, "b", 2, 6
    checkError E_INVALIDINPUT, "Insert() fails on odd number of arguments with keys and values given directly as arguments"
    On Error GoTo 0
    On Error Resume Next
    Dim l1 As List: Set l = List_create("c", 3, "d", 4, 6)
    m1.insert l1
    checkError E_INVALIDINPUT, "Insert() fails on odd number of arguments with keys and values given as single list"
    On Error GoTo 0
    On Error Resume Next
    Dim k1 As List: Set k1 = List_create("e", "f")
    Dim v1 As List: Set v1 = List_create(5, 6, 7)
    m1.insert k1, v1
    checkError E_INVALIDINPUT, "Insert() fails on odd number of arguments with keys and values given as separate lists"
    On Error GoTo 0
    On Error Resume Next
    Dim a1 As Variant: a1 = Array("g", 7, "h", 8, 10)
    m1.insert a1
    checkError E_INVALIDINPUT, "Insert() fails on odd number of arguments with keys and values given as single array"
    On Error GoTo 0
    On Error Resume Next
    Dim kArr1 As Variant: kArr1 = Array("i", "j")
    Dim vArr1 As Variant: vArr1 = Array(9, 10, 11)
    m1.insert kArr1, vArr1
    checkError E_INVALIDINPUT, "Insert() fails on odd number of arguments with keys and values given as separate arrays"
    On Error GoTo 0
    
    gStop
End Sub

Private Sub test_clone()
    gStart "clone"
    
    Dim m As Map: Set m = Map_create
    m("a") = 1
    m("b") = 2
    m("c") = 3
    Dim m2 As Map: Set m2 = m.clone
    equals m2("b"), 2, "clone() works as expected"
    
    gStop
End Sub

Private Sub test_keyvaltypename()
    gStart "keyvaltypename"
    
    Dim m As Map: Set m = Map_create
    equals m.keyTypeName, vbNullString, "keyTypeName() returns null string for untyped Map"
    Dim m2 As Map: Set m2 = Map_createT("Integer", "String")
    equals m2.keyTypeName, "Integer", "keyTypeName() returns the right key type"
    equals m2.valTypeName, "String", "valTypeName() returns the right value type"
    
    gStop
End Sub

Private Sub test_equals()
    gStart "equals"
    
    Dim m As Map: Set m = Map_create
    Dim m2 As Map: Set m2 = Map_create
    m("a") = 1
    m("sdf") = 23
    m2("a") = 1
    m2("sdf") = 23
    equals m, m2, "equals() works with different Maps that have the same keys and values"
    m2("b") = 2
    equals m.equals(m2), False, "equals() returns False when Maps have different keys/values"
    Set m("b") = m2
    Set m2("b") = m2
    equals m, m2, "equals() works with different Maps that have values which are the same object"
    
    Dim m3 As Map: Set m3 = Map_create
    Dim m4 As Map: Set m4 = Map_createT("Integer", "String")
    m3(1) = "a"
    m4(1) = "a"
    equals m3.equals(m4), False, "equals() returns False when Maps have the same values but different type constraints"
    
    gStop
End Sub
