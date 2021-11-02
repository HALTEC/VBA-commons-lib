Attribute VB_Name = "Test_Setx"
Option Explicit

Private Sub test_all()
    gStart "Setx"
    
    test_create
    test_typed
    test_innerTypeName
    test_elems
    test_add
    test_append
    test_remove
    test_removeall
    test_exists
    test_equals
    test_discard
    test_pick
    test_pop
    test_values
    test_for_each
    test_union
    test_intersection
    test_difference
    test_symmetric_difference
    test_is_disjoint
    test_is_subset
    test_is_superset
    test_update_union
    test_update_intersection
    test_update_difference
    test_update_symmetric_difference
    test_count
    
    gStop
End Sub

Private Sub test_create()
    gStart "create"
    
    pass "Alles toll"
    Dim s As Setx: Set s = Setx_create
    gStop
End Sub

Private Sub test_typed()
    gStart "typed"
    
    Dim s As Setx: Set s = Setx_create
    equals s.typed, False, "Untyped set says so"
    
    Dim s2 As Setx: Set s2 = Setx_createLT
    equals s2.typed, True, "Lazy typed set says so"
    
    Dim s3 As Setx: Set s3 = Setx_createT("Integer")
    equals s3.typed, True, "Typed set says so"
    
    gStop
End Sub


Private Sub test_innerTypeName()
    gStart "innerTypeName"
    
    Dim s As Setx: Set s = Setx_create
    equals s.innerTypeName, "", "Untyped set has typename of """""
    
    Set s = Setx_createLT
    equals s.innerTypeName, "", "Empty lazy typed set has typename of """""
    
    Set s = Setx_createLT(1)
    equals s.innerTypeName, "Integer", "Non-empty lazy typed set has correct typename"
    
    Set s = Setx_createT("Integer")
    equals s.innerTypeName, "Integer", "Typed set has correct typename"
    
    Dim untypedVar As Variant
    Set s = Setx_createLT(untypedVar)
    equals s.innerTypeName, "", "Adding an empty Variant does not set typename"
    
    Dim intTypedVar As Variant: intTypedVar = 5
    Set s = Setx_createLT(intTypedVar)
    equals s.innerTypeName, "Integer", "Adding an int Variant sets the type name to the inner type"
    
    Dim setTypedVar As Variant: Set setTypedVar = Setx_create
    Set s = Setx_createLT
    s.add setTypedVar
    equals s.innerTypeName, "Setx", "Adding a set object Variant sets the type name to the inner type"
    
    gStop
End Sub

Private Sub test_elems()
    gStart "elems"
    
    Dim s As Setx: Set s = Setx_create
    equals s.elems, 0, "elems works for empty set"
    
    gStop
End Sub

Private Sub test_add()
    gStart "add"
    
    Dim s As Setx: Set s = Setx_create
    s.add 20
    equals s.elems, 1, "Set contains one element"
    
    s.add 20
    equals s.elems, 1, "Adding existing element shouldn't duplicate"
    
    gStop

End Sub

Private Sub test_append()
    
    gStart "append"
    
    Dim s As Setx: Set s = Setx_create
    s.append List_create(1, 2, 3)
    equals s.elems, 3, "Append can add multiple elements via a list"
    
    Set s = Setx_create
    Dim arr(2) As Integer
    arr(0) = 1
    arr(1) = 2
    arr(2) = 3
    s.append arr
    equals s.elems, 3, "Append can add multiple elements via an array"
    
    Set s = Setx_create
    s.append List_create
    equals s.elems, 0, "Appending an empty list adds no elements"
    
    Set s = Setx_create
    Dim arr2() As Integer
    arr2 = Arrays.emptyIntegerArray
    s.append arr2
    equals s.elems, 0, "Appending an empty array adds no elements"
    
    Set s = Setx_create
    Dim other As Setx: Set other = Setx_create(1, 2, 3)
    s.append other
    equals s.elems, 3, "Append can add multiple elements via a set"
    
    gStop
End Sub

Private Sub test_remove()
    gStart "remove"
    
    Dim s As Setx: Set s = Setx_create(0, 1)
    equals s.elems, 2, "Remove() Setup (1)"
    
    s.remove 1
    equals s.elems, 1, "Remove() removes one element from set (2)"
    equals s.exists(2), False, "Remove() removes correct element (3)"
    
    On Error Resume Next
    s.remove 3
    checkError E_ARGUMENTOUTOFRANGE, "Remove fails on missing element"
    On Error GoTo 0
    
    gStop
End Sub

Private Sub test_removeall()
    gStart "removeall"
    
    Dim s As Setx: Set s = Setx_create(0, 1)
    equals s.elems, 2, "RemoveAll() Setup (1)"
    
    s.removeAll
    equals s.elems, 0, "RemoveAll() actually removes all elements (2)"
    
    On Error Resume Next
    s.removeAll
    checkNoError "Calling RemoveAll() doesn't cause error message"
    On Error GoTo 0
    
    gStop
End Sub

Private Sub test_exists()
    gStart "exists"
    
    Dim s As Setx: Set s = Setx_create(3, 100)
    equals s.exists(100), True, "Exists() returns True if key exists (1)"
    equals s.exists(3.14), False, "Exists() returns False if key doesn't exist (2)"
    
    gStop
End Sub

Private Sub test_equals()

    gStart "equals"
    
    Dim s As Setx: Set s = Setx_create
    Dim other As Setx: Set other = Setx_create
    equals s.equals(other), True, "Equals() deems two empty sets as identical (1)"
    
    Set s = Setx_create(1, 2, 3)
    Set other = Setx_create(2, 3, 4)
    equals s.equals(other), False, "Equals() works as expected when sets aren't identical (2)"
    
    Set other = Setx_create(1, 2, 3)
    equals s.equals(other), True, "Equals() works as expected when sets are identical (3)"
    
    gStop
End Sub

Private Sub test_discard()
    gStart "discard"
    
    Dim s As Setx: Set s = Setx_create(0, 1)
    equals s.elems, 2, "Discard() Setup (1)"
    
    s.remove 0
    equals s.elems, 1, "Discard() Setup (2)"
    equals s.exists(2), False, "Discard() Setup(3)"
    
    On Error Resume Next
    s.discard 3
    checkNoError "Calling Discard() doesn't cause an error message if to be removed element isn't in set (4)"
    On Error GoTo 0
    
    gStop
End Sub

Private Sub test_pick()

    gStart "pick"
    
    Dim s As Setx: Set s = Setx_create("a", "b", "c")
    Dim pick As Variant
    Dim x As Integer
    Dim a_seen As Boolean
    Dim b_seen As Boolean
    Dim c_seen As Boolean
    For x = 0 To 32767
        Dim choice As String
        choice = s.pick
        If choice = "a" Then
            a_seen = True
        ElseIf choice = "b" Then
            b_seen = True
        ElseIf choice = "c" Then
            c_seen = True
        End If
        If a_seen = True And b_seen = True And c_seen = True Then
            Exit For
        End If
    Next
    
    equals a_seen, True, "Pick() doesn't exclude a (1)"
    equals b_seen, True, "Pick() doesn't exclude b (2)"
    equals c_seen, True, "Pick() doesn't exclude c (3)"
    
    gStop
End Sub

Private Sub test_pop()
    gStart "pop"
    
    Dim s As Setx: Set s = Setx_create(0, 1, 2)
    s.pop
    equals s.elems, 2, "Pop() removes unspecified element from set (1)"
    
    s.pop
    equals s.elems, 1, "Pop() removes another element from set (2)"
    
    s.pop
    equals s.elems, 0, "Pop() removes another element from set (3)"
    
    On Error Resume Next
    s.pop
    checkNoError "Calling Pop() doesn't cause an error message if called on an empty set (4)"
    On Error GoTo 0
    
    gStop
End Sub

Public Sub test_values()

    gStart "values"
    
    Dim s As Setx: Set s = Setx_create(1, 2)
    Dim values As List
    Set values = s.values
    equals values.elems, 2, "Values() creates list that contains the same number of elements as the underlying set (1)"
    equals values.contains(2), True, "Values() creates list that contains all elements of the underlying set (2)"
    
    gStop
End Sub


Public Sub test_for_each()
    
    gStart "For Each"
    
    Dim s As Setx: Set s = Setx_create(1)
    Dim t As Variant
    For Each t In s
        equals t, 1, "Iteration works"
    Next
    
    gStop
End Sub

Public Sub test_union()
    
    gStart "Union"
    
    Dim s As Setx: Set s = Setx_create(1, 2, 3)
    Dim other As Setx: Set other = Setx_create(3, 4, 5)
    Dim union As Setx: Set union = s.union(other)
    equals union.elems, 5, "Union() creates set with the correct amount of elements without duplications (1)"
    equals union.exists(1), True, "Union() creates set containing elements from Me(2)"
    equals union.exists(5), True, "Union() creates set containing elements from other (3)"
    
    Set s = Setx_create(1)
    Set other = Setx_create
    Set union = s.union(other)
    equals union.elems, 1, "Union() contains one element when called on a unielemental Me and empty set(4)"
    
    Set s = Setx_create
    Set other = Setx_create(1)
    Set union = s.union(other)
    equals union.elems, 1, "Union() contains one element when called on an empty Me and a unielemental set(5)"
    
    Set s = Setx_create
    Set other = Setx_create
    Set union = s.union(other)
    equals union.elems, 0, "Union () is empty when Me and other are empty(6)"
    
    gStop
End Sub

Public Sub test_intersection()

    gStart "Intersection"
    
    Dim s As Setx: Set s = Setx_create(1, 2, 3)
    Dim other As Setx: Set other = Setx_create(2)
    Dim intersection As Setx: Set intersection = s.intersection(other)
    equals intersection.elems, 1, "Intersection() contains one element when there is only one common element between Me and other (1)"
    equals intersection.exists(2), True, "Intersection() contains the one common element between Me and other (2)"
    
    Set s = Setx_create(1)
    Set other = Setx_create
    Set intersection = s.intersection(other)
    equals intersection.elems, 0, "Intersection() is empty when Me has one element and other is empty (3)"
    
    Set s = Setx_create
    other.add 1
    Set intersection = s.intersection(other)
    equals intersection.elems, 0, "Intersection() is empty when Me is empty and other has one element (4)"
    
    Set s = Setx_create
    Set other = Setx_create
    Set intersection = s.intersection(other)
    equals intersection.elems, 0, "Intersection() works as expected (5)"
    
    Set s = Setx_create(1)
    Set other = Setx_create(2)
    Set intersection = s.intersection(other)
    equals intersection.elems, 0, "Intersection() is empty when Me and other have no common elements (6)"
    
    gStop
    
End Sub

Public Sub test_difference()

    gStart "Difference"
    
    Dim s As Setx: Set s = Setx_create(1, 2, 3)
    Dim other As Setx: Set other = Setx_create(2)
    Dim difference As Setx: Set difference = s.difference(other)
    equals difference.elems, 2, "The number of elements of the set that Difference() creates correlates to the amount of the difference between Me and other(1)"
    equals difference.exists(3), True, "Difference() returns set with the actual elements that Me has that other doesn't (2)"
    
    Set s = Setx_create
    Set other = Setx_create
    Set difference = s.difference(other)
    equals difference.elems, 0, "Difference() returns empty set when Me and other are empty (3)"
    
    Set s = Setx_create(1)
    Set other = Setx_create
    Set difference = s.difference(other)
    equals difference.elems, 1, "Difference() returns set with one element when Me has one element and other is empty (4)"
    equals difference.exists(1), True, "Difference() returns set containing the element Me has that empty other doesn't (5)"
    
    Set difference = other.difference(s)
    equals difference.elems, 0, "Difference() returns empty set when Me is empty and other has one element (6)"

    Set s = Setx_create(1)
    Set other = Setx_create(1)
    Set difference = s.difference(other)
    equals difference.elems, 0, "Difference() returns empty set when Me and other are identical (7)"
    
    gStop
End Sub

Public Sub test_symmetric_difference()

    gStart "Symmetric-Difference"
    
    Dim s As Setx: Set s = Setx_create(1, 2, 3)
    Dim other As Setx: Set other = Setx_create(3, 4, 5)
    Dim symmetric_difference As Setx: Set symmetric_difference = s.symmetric_difference(other)
    equals symmetric_difference.elems, 4, "Symmetric-Difference() returns set with an amount of elements that correlates to the amount of elements found in Me or other but not in both (1)"
    equals symmetric_difference.exists(2), True, "Symmetric-Difference() returns set containing all elements found in Me and not in other (2)"
    equals symmetric_difference.exists(3), False, "Symmetric-Difference() returns set not containing all elements found in Me and other (3)"
    equals symmetric_difference.exists(4), True, "Symmetric-Difference() returns set containing all elements found in other and not in Me(4)"
    
    Set s = Setx_create
    Set other = Setx_create
    Set symmetric_difference = s.symmetric_difference(other)
    equals symmetric_difference.elems, 0, "Symmetric-Difference() returns empty set when Me and other are empty (5)"
    
    Set s = Setx_create(1)
    Set other = Setx_create
    Set symmetric_difference = s.symmetric_difference(other)
    equals symmetric_difference.elems, 1, "Symmetric-Difference() returns set with the same amount of elements as Me if other is empty and vice versa (6)"
    equals symmetric_difference.exists(1), True, "Symmetric-Difference() returns set with the  same elements as Me if other is empty and vice versa (7)"
    
    gStop
    
End Sub

Public Sub test_is_disjoint()

    gStart "Is disjoint"
    
    Dim s As Setx: Set s = Setx_create
    Dim other As Setx: Set other = Setx_create
    equals s.is_disjoint(other), True, "Is disjoint() returns True when Me and other are empty (1)"
    equals s.is_disjoint(s), True, "Is disjoint() returns True when calling it on a set and passing the identical set as other (2)"
    
    s.add 1
    other.add 2
    equals s.is_disjoint(other), True, "Is disjoint() returns True when both Me and other have a certain amount of elements but contain no common elements (3)"
    
    s.add 2
    equals s.is_disjoint(other), False, "Is disjoint() returns False when at least one element is common between Me and other (4)"
    
    gStop
    
End Sub

Public Sub test_is_subset()

    gStart "Is subset"
    
    Dim s As Setx: Set s = Setx_create
    Dim other As Setx: Set other = Setx_create
    equals s.is_subset(other), True, "Is subset() returns True when both sets are empty (1)"
    
    s.add 1
    other.add 1
    equals s.is_subset(other), True, "Is subset() returns True when both sets are identical (2)"
    
    other.add 2
    equals s.is_subset(other), True, "Is subset() returns True when all elements in Me are contained in other (3)"
    
    s.add 2
    
    Set other = Setx_create
    equals s.is_subset(other), False, "Is subset() returns False when Me contains elements that other doesn't (4)"
    
    gStop
End Sub

Public Sub test_is_superset()

    gStart "Is superset"
    
    Dim s As Setx: Set s = Setx_create
    Dim other As Setx: Set other = Setx_create
    equals s.is_superset(other), True, "Is superset() returns True when both sets are empty (1)"
    
    s.add 1
    equals s.is_superset(other), True, "Is superset() returns True when all is empty (2)"
    
    other.add 1
    equals s.is_superset(other), True, "Is superset() returns True when both sets are identical (3)"
    
    other.add 2
    equals s.is_superset(other), False, "Is superset() returns False when other contains elements that Me doesn't (4)"
    
    gStop
End Sub

Public Sub test_update_union()

    gStart "Update Union"
    
    Dim s As Setx: Set s = Setx_create(1, 2, 3)
    Dim other As Setx: Set other = Setx_create(3, 4, 5)
    s.update_union other
    equals s.elems, 5, "Update Union() makes set end up with correct amount of elements (1)"
    equals s.exists(1), True, "Update Union() makes set contain all elements from Me (2)"
    equals s.exists(5), True, "Update Union() makes set contain all elements from other (3)"
    
    gStop
End Sub

Public Sub test_update_intersection()

    gStart "Update Intersection"
    
    Dim s As Setx: Set s = Setx_create(1, 2, 3)
    Dim other As Setx: Set other = Setx_create(3, 4, 5)
    s.update_intersection other
    equals s.elems, 1, "Update Intersection() modifies set, making it have the correct amount of elements (1)"
    equals s.exists(3), True, "Update Intersection() makes set contain all elements that are both in Me and in other (2)"
    
    Set other = Setx_create
    s.update_intersection other
    equals s.elems, 0, "Update Intersection() modifies set to an empty set in the abscense of common elements(3)"
    gStop
End Sub

Public Sub test_update_difference()

    gStart "Update Difference"
    
    Dim s As Setx: Set s = Setx_create
    Dim other As Setx: Set other = Setx_create
    s.update_difference other
    equals s.elems, 0, "Update Difference() modifies s to an empty set when Me and other are empty sets (1)"
    
    Set s = Setx_create(1, 2, 3)
    s.update_difference other
    equals s.elems, 3, "Update Difference() modifies s so that s contains all elements in Me that aren't in other (2)"
    
    Set s = Setx_create(1, 2, 3)
    Set other = Setx_create(3, 4)
    s.update_difference other
    equals s.elems, 2, "Update Difference() modifies s so that s contains all elements in Me that aren't in other (3)"
    
    Set s = Setx_create(1, 2, 3)
    s.update_difference s
    equals s.elems, 0, "Update Difference() modifies s so that s is empty when examining to difference with identical set (4)"
    
    gStop
End Sub

Public Sub test_update_symmetric_difference()
    gStart "Update Symmetric Difference"
    
    Dim s As Setx: Set s = Setx_create
    s.update_symmetric_difference s
    equals s.elems, 0, "Update Symmetric Difference() modifies s so that s us empty when examining symmetric difference to itself (1)"
    
    Set s = Setx_create(1, 2, 3)
    Dim other As Setx: Set other = Setx_create(3, 4, 5)
    s.update_symmetric_difference other
    equals s.elems, 4, "Update Symmetric Difference() modifies s so that s contains all elements of M that aren't in other and all elements of other that aren't in Me (2)"
    
    gStop
End Sub

Public Sub test_count()

    gStart "Count"
    
    Dim s As Setx: Set s = Setx_create(1, 2)
    s.count
    equals s.elems, 2, "Count() counts the amount of elements in a set correctly"

    gStop
    
End Sub
