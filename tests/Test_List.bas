Attribute VB_Name = "Test_List"
Option Explicit

Private Sub test_all()
    gStart "List"
    
    test_create
    test_createT
    test_createLT
    test_item
    test_forEach
    test_elems
    test_typed
    test_innerTypeName
    test_pushPop
    test_append
    test_shiftUnshift
    test_prepend
    test_splice
    test_clone
    test_getRange
    test_remove
    test_firstLast
    test_clear
    test_contains
    test_indexOfLastIndexOf
    test_maxMin
    test_reverse
    test_sort
    test_uniq
    test_equals
    
    gStop
End Sub

Private Sub test_create()
    gStart "create"
    
    On Error Resume Next
    List_create
    checkNoError "Creating an empty untyped list"
    On Error GoTo 0
    
    Dim l As List: Set l = List_create
    equals l.elems, 0, "Empty untyped list is empty"

    On Error Resume Next
    List_create 1, 2, 3
    checkNoError "Creating a non-empty untyped list"
    On Error GoTo 0
    
    Dim l2 As List: Set l2 = List_create(1, 2, 3)
    equals l2.elems, 3, "1 Filling untyped list on initialize works"
    equals l2(0), 1, "2 Filling untyped list on initialize works"

    On Error Resume Next
    List_create "Hey", 1, 2
    checkNoError "Creating a mixed type list"
    On Error GoTo 0

    gStop
End Sub

Private Sub test_createT()
    gStart "createT"
    
    On Error Resume Next
    List_createT "Integer"
    checkNoError "Creating an empty typed list"
    On Error GoTo 0
    
    Dim l As List: Set l = List_createT("Integer")
    equals l.elems, 0, "Empty typed list is empty"

    Dim l2 As List: Set l2 = List_createT("Integer")
    l2.append 1, 2, 3
    equals l2.elems, 3, "1 Filling typed list works"
    equals l2(0), 1, "2 Filling typed list works"

    
    Dim l3 As List: Set l3 = List_createT("Integer")
    
    On Error Resume Next
    l3.push "Hey"
    checkError E_TYPEMISMATCH, "Typed list fails on inserting wrong type"
    On Error GoTo 0

    gStop
End Sub

Private Sub test_createLT()
    gStart "create"
    
    On Error Resume Next
    List_createLT
    checkNoError "Creating an empty lazy typed list"
    On Error GoTo 0
    
    Dim l As List: Set l = List_createLT
    equals l.elems, 0, "Empty lazy typed list is empty"

    On Error Resume Next
    List_createLT 1, 2, 3
    checkNoError "Creating a non-empty lazy typed list"
    On Error GoTo 0
    
    Dim l2 As List: Set l2 = List_createLT(1, 2, 3)
    equals l2.elems, 3, "1 Filling lazy typed list on initialize works"
    equals l2(0), 1, "2 Filling lazy typed list on initialize works"

    Dim l3 As List: Set l3 = List_createLT(1)
    
    On Error Resume Next
    l3.push "Hey"
    checkError E_TYPEMISMATCH, "Lazy typed list fails on inserting wrong type"
    On Error GoTo 0

    gStop
End Sub

Private Sub test_item()
    gStart "item"
    
    Dim dummy As Variant
    
    Dim l As List
    Set l = List_create(1, 2, 3)
    
    equals l(0), 1, "1 Item access (0-based)"
    equals l(1), 2, "2 Item access (0-based)"
    equals l(2), 3, "3 Item access (0-based)"
    
    On Error Resume Next
    dummy = l(3)
    checkError E_INDEXOUTOFRANGE, "Accessing index >= count fails"
    On Error GoTo 0

    On Error Resume Next
    dummy = l(-4)
    checkError E_INDEXOUTOFRANGE, "Accessing index < -count fails"
    On Error GoTo 0
    
    On Error Resume Next
    l(0) = 4
    checkNoError "Item set doesn't error"
    On Error GoTo 0
    
    equals l(0), 4, "Item set actually sets"
    equals l.elems, 3, "Item set does not change size"

    gStop
End Sub

Private Sub test_forEach()
    gStart "forEach"
    
    Dim l As List: Set l = List_createLT(1, 2, 3)
    Dim runner As Variant
    
    On Error Resume Next
    For Each runner In l
    Next
    checkNoError "For Each does not throw"
    On Error GoTo 0
    
    Dim counter As Integer: counter = 1
    For Each runner In l
        equals runner, counter, counter & " For Each iterates correctly"
        counter = counter + 1
    Next
    
    gStop
End Sub

Private Sub test_elems()
    gStart "elems"
    
    Dim l As List: Set l = List_createLT
    equals l.elems, 0, "elems works for empty list"
    
    Dim l2 As List: Set l2 = List_createLT(1, 2, 3)
    equals l2.elems, 3, "elems works for non empty list"
    
    gStop
End Sub

Private Sub test_typed()
    gStart "typed"
    
    Dim l As List: Set l = List_create
    equals l.typed, False, "Untyped list says so"
    
    Dim l2 As List: Set l2 = List_createLT
    equals l2.typed, True, "Lazy typed list says so"
    
    Dim l3 As List: Set l3 = List_createT("Integer")
    equals l3.typed, True, "Typed list says so"
    
    gStop
End Sub

Private Sub test_innerTypeName()
    gStart "innerTypeName"
    
    Dim l As List
    
    Set l = List_create
    equals l.innerTypeName, "", "Untyped list has typename of """""
    
    Set l = List_createLT
    equals l.innerTypeName, "", "Empty lazy typed list has typename of """""
    
    Set l = List_createLT(1)
    equals l.innerTypeName, "Integer", "Non-empty lazy typed list has correct typename"
    
    Set l = List_createT("Integer")
    equals l.innerTypeName, "Integer", "Typed list has correct typename"
    
    Dim arr() As Integer
    Set l = List_createLT(arr)
    equals l.innerTypeName, "Integer()", "Adding an array sets the typename to array"
    
    Dim untypedVar As Variant
    Set l = List_createLT(untypedVar)
    equals l.innerTypeName, "", "Adding an empty Variant does not set typename"
    
    Dim intTypedVar As Variant: intTypedVar = 5
    Set l = List_createLT(intTypedVar)
    equals l.innerTypeName, "Integer", "Adding an int Variant sets the type name to the inner type"
    
    Dim listTypedVar As Variant: Set listTypedVar = List_create
    Set l = List_createLT(listTypedVar)
    equals l.innerTypeName, "List", "Adding a list object Variant sets the type name to the inner type"
    
    Dim arr2() As Integer
    Dim arrTypedVar As Variant: arrTypedVar = arr2
    Set l = List_createLT(arrTypedVar)
    equals l.innerTypeName, "Integer()", "Adding an array Variant sets the typename to array"
    
    gStop
End Sub

Private Sub test_pushPop()
    gStart "push & pop"
    
    Dim l As List
    
    ' push
    Set l = List_create(1)
    
    l.push 2
    
    equals l.elems, 2, "push adds an element"
    equals l(1), 2, "push adds the element at the end"
    
    Set l = List_create(1)
    Dim arr(2) As Integer
    arr(0) = 2
    arr(1) = 3
    l.push arr
    
    equals l.elems, 2, "Pushing an array does not unpack"
    equals l(1), arr, "Pushing an array acutally adds the array"
    
    ' pop
    Set l = List_create(1, 2, 3)
    
    equals l.pop, 3, "pop returns the last element"
    equals l.elems, 2, "pop removes the last element"
    
    Set l = List_create
    On Error Resume Next
    l.pop
    checkError E_ILLEGALSTATE, "Calling pop on an empty list throws E_ILLEGALSTATE"
    On Error GoTo 0
    
    gStop
End Sub

Private Sub test_append()
    gStart "append"
    
    Dim l As List
    
    Set l = List_create(1)
    l.append 2, 3, 4
    
    equals l.elems, 4, "Append can add multiple elements"
    equals l(1), 2, "1 Append adds the elements at the right"
    equals l(2), 3, "2 Append adds the elements at the right"
    equals l(3), 4, "3 Append adds the elements at the right"
    
    Set l = List_create
    l.append List_create(1, 2, 3)
    equals l.elems, 3, "Append can add multiple elements via a list"
    equals l(0), 1, "1 Append adds the right elements via a list"
    equals l(1), 2, "2 Append adds the right elements via a list"
    equals l(2), 3, "3 Append adds the right elements via a list"
    
    Set l = List_create
    Dim arr(2) As Integer
    arr(0) = 1
    arr(1) = 2
    arr(2) = 3
    l.append arr
    equals l.elems, 3, "Append can add multiple elements via an array"
    equals l(0), 1, "1 Append adds the right elements via an array"
    equals l(1), 2, "2 Append adds the right elements via an array"
    equals l(2), 3, "3 Append adds the right elements via an array"
    
    Set l = List_create
    l.append List_create
    equals l.elems, 0, "Appending an empty list adds no elements"
    
    Set l = List_create
    Dim arr2() As Integer
    arr2 = Arrays.emptyIntegerArray
    l.append arr2
    equals l.elems, 0, "Appending an empty array adds no elements"
    
    gStop
End Sub

Private Sub test_shiftUnshift()
    gStart "shift & unshift"
    
    Dim l As List
    
    ' unshift
    Set l = List_create(2)
    
    l.unshift 1
    
    equals l.elems, 2, "unshift adds an element"
    equals l(0), 1, "unshift adds the element at the front"
    
    Set l = List_create(3)
    Dim arr(2) As Integer
    arr(0) = 1
    arr(1) = 2
    l.unshift arr
    
    equals l.elems, 2, "Unshifting an array does not unpack"
    equals l(0), arr, "Unshifting an array acutally adds the array"
    
    ' shift
    Set l = List_create(1, 2, 3)
    
    equals l.shift, 1, "shift returns the first element"
    equals l.elems, 2, "shift removes the first element"
    
    Set l = List_create
    On Error Resume Next
    l.shift
    checkError E_ILLEGALSTATE, "Calling shift on an empty list throws E_ILLEGALSTATE"
    On Error GoTo 0
    
    gStop
End Sub

Private Sub test_prepend()
    gStart "prepend"
    
    Dim l As List
    
    Set l = List_create(1)
    l.prepend 2, 3, 4
    
    equals l.elems, 4, "prepend can add multiple elements"
    equals l(0), 2, "1 prepend adds the elements at the right"
    equals l(1), 3, "2 prepend adds the elements at the right"
    equals l(2), 4, "3 prepend adds the elements at the right"
    equals l(3), 1, "4 prepend adds the elements at the right"
    
    Set l = List_create
    l.prepend List_create(1, 2, 3)
    equals l.elems, 3, "prepend can add multiple elements via a list"
    equals l(0), 1, "1 prepend adds the right elements via a list"
    equals l(1), 2, "2 prepend adds the right elements via a list"
    equals l(2), 3, "3 prepend adds the right elements via a list"
    
    Set l = List_create
    Dim arr(2) As Integer
    arr(0) = 1
    arr(1) = 2
    arr(2) = 3
    l.prepend arr
    equals l.elems, 3, "prepend can add multiple elements via an array"
    equals l(0), 1, "1 prepend adds the right elements via an array"
    equals l(1), 2, "2 prepend adds the right elements via an array"
    equals l(2), 3, "3 prepend adds the right elements via an array"
    
    Set l = List_create
    l.prepend List_create
    equals l.elems, 0, "Prepending an empty list adds no elements"
    
    Set l = List_create
    Dim arr2() As Integer
    arr2 = Arrays.emptyIntegerArray
    l.prepend arr2
    equals l.elems, 0, "Prepending an empty array adds no elements"
    
    gStop
End Sub

Private Sub test_splice()
    gStart "splice"
    
    Dim l As List
    
    Set l = List_create(1, 2, 3, 4, 5)
    l.splice 1, 2, List_create("two", "three")
    equals l, List_create(1, "two", "three", 4, 5), "simple splice"
    
    
    Set l = List_create(1, 2, 3, 4, 5)
    l.splice 1, 2
    equals l, List_create(1, 4, 5), "without replacement"
        
        
    Set l = List_create(1, 2, 3, 4, 5)
    l.splice 1, 2, List_create
    equals l, List_create(1, 4, 5), "with empty list as replacement"
    
    
    Set l = List_create(1, 2, 3, 4, 5)
    l.splice 1, 0, List_create("hi", "ho")
    equals l, List_create(1, "hi", "ho", 2, 3, 4, 5), "without removals"
    
    
    Set l = List_create(1, 2, 3, 4, 5)
    l.splice 0, 0, List_create("hi", "ho")
    equals l, List_create("hi", "ho", 1, 2, 3, 4, 5), "without removals before first element"
    
    
    Set l = List_create(1, 2, 3, 4, 5)
    l.splice 5, 0, List_create("hi", "ho")
    equals l, List_create(1, 2, 3, 4, 5, "hi", "ho"), "without removals after last element"
    
    
    Set l = List_create(1, 2, 3)
    
    On Error Resume Next
    l.splice 1, 3
    checkError E_INDEXOUTOFRANGE, "Splicing beyond the end fails"
    On Error GoTo 0
    
    
    Set l = List_create(1, 2, 3)
    
    On Error Resume Next
    l.splice 1, 2
    checkNoError "Splicing up to the end succeeds"
    On Error GoTo 0
    
    equals l, List_create(1), "Splicing to the end is correct"
    
    
    Set l = List_create(1, 2, 3)
    l.splice -1, 1, List_create("last")
    equals l, List_create(1, 2, "last"), "Negative indexes work"
    
    gStop
End Sub

Private Sub test_clone()
    gStart "clone"
    
    Dim l As List
    Dim l2 As List
    
    Set l = List_create(1, 2, 3)
    Set l2 = l.clone
    
    equals l, l2, "cloned list equals"
    
    l2.push 4
    
    equals l.elems, 3, "cloned list is  a clone and not a reference"
    equals l2.elems, 4, "cloned list is  a clone and not a reference"
    
    
    Set l = List_create(1, 2, 3, List_create(4, 5))
    Set l2 = l.clone
    
    equals l, l2, "cloned list with sublist equals"
    
    equals l(3).elems, 2, "1 cloned list is a flat clone"
    l2(3).push 6
    equals l(3).elems, 3, "2 cloned list is a flat clone"
    
    gStop
End Sub

Private Sub test_getRange()
    gStart "getRange"
    
    Dim l As List
    Set l = List_create(1, 2, 3)
    
    equals l.getRange(0, 1), List_create(1)
    equals l.getRange(0, 0), List_create
    equals l.getRange(2, 1), List_create(3)
    equals l.getRange(0, 3), List_create(1, 2, 3)
    
    gStop
End Sub

Private Sub test_remove()
    gStart "remove"
    
    Dim l As List
    
    Set l = List_create(1, 2, 3)
    l.remove 0
    equals l, List_create(2, 3), "Remove single element"
    
    Set l = List_create(1, 2, 3)
    l.remove 0, 2
    equals l, List_create(3), "Remove multiple elements"
    
    Set l = List_create(1, 2, 3)
    l.remove 0, 0
    equals l, List_create(1, 2, 3), "Remove no elements"
    
    Set l = List_create(1, 2, 3)
    l.remove 0, 3
    equals l, List_create, "Remove all elements"
    
    Set l = List_create(1, 2, 3)
    l.remove 2, 1
    equals l, List_create(1, 2), "Remove an element not at the start"
    
    gStop
End Sub

Private Sub test_firstLast()
    gStart "first & last"
    
    Dim l As List
    
    Set l = List_create(1, 2, 3)
    equals l.first, 1, "first returns first element"
    equals l.last, 3, "last returns last element"
    
    Set l = List_create
    
    On Error Resume Next
    l.first
    checkError E_ILLEGALSTATE, "first throws when list is empty"
    On Error GoTo 0
    
    On Error Resume Next
    l.last
    checkError E_ILLEGALSTATE, "last throws when list is empty"
    On Error GoTo 0

    gStop
End Sub

Private Sub test_clear()
    gStart "clear"
    
    Dim l As List
    
    Set l = List_create(1, 2, 3)
    l.clear
    equals l.elems, 0, "non empty list"
    
    Set l = List_create
    l.clear
    equals l.elems, 0, "empty list"
    
    gStop
End Sub

Private Sub test_contains()
    gStart "contains"
    
    Dim l As List
    
    Set l = List_create(1, 2, 3)
    equals l.contains(2), True, "element is contained"
    equals l.contains(5), False, "element is not contained"
    
    Dim obj As New Math
    Dim obj2 As New Math
    
    Set l = List_create(obj, 1, 2)
    equals l.contains(obj), True, "Object element is contained"
    equals l.contains(obj2), False, "Object element is not contained"
    
    gStop
End Sub

Private Sub test_indexOfLastIndexOf()
    gStart "indexOf & lastIndexOf"
    
    Dim l As List: Set l = List_create(1, 2, 3, 1, 2, 3)
    
    equals l.indexOf(2), 1, "contained element"
    equals l.indexOf(4), -1, "not contained element"
    equals l.lastIndexOf(2), 4, "contained element"
    equals l.lastIndexOf(4), -1, "not contained element"
    
    gStop
End Sub

Private Sub test_maxMin()
    gStart "max & min"
    
    Dim l As List: Set l = List_create(1, 2, 3, 1, 2, 3)
    
    equals l.min, 1, "min"
    equals l.max, 3, "max"
    
    gStop
End Sub

Private Sub test_reverse()
    gStart "reverse"
    
    equals List_create(1, 2, 3).reverse, List_create(3, 2, 1)
    
    gStop
End Sub

Private Sub test_sort()
    gStart "sort"
    
    equals List_createLT(4, 2, 5, 3, 1, 6, 7, 5, 6).sort, List_createLT(1, 2, 3, 4, 5, 5, 6, 6, 7)
    equals List_createLT("Hi", "there", "Hi", "me").sort, List_createLT("Hi", "Hi", "me", "there")
    equals List_create().sort, List_create()
    
    On Error Resume Next
    List_create(Strings).sort
    checkError E_ILLEGALSTATE, "Sorting unsortable things fails"
    On Error GoTo 0
    
    gStop
End Sub

Private Sub test_uniq()
    gStart "uniq"
    
    equals List_create(1, 1, 2, 1, 3).uniq, List_create(1, 2, 3)
    equals List_create("a", "b", "c", "b").uniq, List_create("a", "b", "c")
    equals List_create().uniq, List_create()
    
    gStop
End Sub

Private Sub test_equals()
    gStart "equals"
    
    ok List_create("asdf", "qwer", "1").equals(List_create("asdf", "qwer", "1"))
    ok Not List_create("asdf", "qwer", "1").equals(List_create("asdf", "qwer", "2"))
    ok Not List_create("asdf", "qwer", "1").equals(List_create("bsdf", "qwer", "1"))
    ok Not List_create("asdf", "qwer", "1").equals(List_create("bsdf", "qwer"))
    ok Not List_create("asdf", "qwer", "1").equals(List_create("bsdf", "qwer", "1", "2"))
    
    gStop
End Sub
