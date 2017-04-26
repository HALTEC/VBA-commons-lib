Attribute VB_Name = "Test_Arrays"
Option Explicit
Option Base 0 ' Default


Private Sub test_all()
    gStart "Arrays"
    
    test_sort
    test_emptyVariantArray
    test_emptyIntegerArray
    test_elems
    test_toVariantArray
    
    gStop
End Sub

Private Sub test_sort()
    gStart "sort"
    
    Dim arr(4) As Integer
    arr(0) = 4
    arr(1) = 2
    arr(2) = 3
    arr(3) = 1
    arr(4) = 0
    
    Arrays.sort arr
    
    equals UBound(arr), 4
    equals arr(0), 0
    equals arr(1), 1
    equals arr(2), 2
    equals arr(3), 3
    equals arr(4), 4
    
    ' Descending
    arr(0) = 4
    arr(1) = 2
    arr(2) = 3
    arr(3) = 1
    arr(4) = 0
    
    Arrays.sort arr, descending
    
    equals UBound(arr), 4
    equals arr(0), 4
    equals arr(1), 3
    equals arr(2), 2
    equals arr(3), 1
    equals arr(4), 0
    
    gStop
End Sub

Private Sub test_emptyVariantArray()
    gStart "emptyVariantArray"
    
    Dim arr() As Variant: arr = Arrays.emptyVariantArray
    equals UBound(arr) - LBound(arr) + 1, 0, "Creates an empty array"
    
    gStop
End Sub

Private Sub test_emptyIntegerArray()
    gStart "emptyIntegerArray"
    
    Dim arr() As Integer: arr = Arrays.emptyIntegerArray
    equals UBound(arr) - LBound(arr) + 1, 0, "Creates an empty array"
    
    gStop
End Sub

Private Sub test_elems()
    gStart "elems"
    
    Dim arr() As Variant: arr = Array()
    equals Arrays.elems(arr), 0, "Empty array"
    
    Dim arr2(4) As Variant
    equals Arrays.elems(arr2), 5, "Static array"
    
    Dim arr3() As Variant
    ReDim arr3(3)
    equals Arrays.elems(arr3), 4, "Dynamic array"
    
    gStop
End Sub

Private Sub test_toVariantArray()
    gStart "toVariantArray"
    
    Dim arr() As Integer
    ReDim arr(3)
    
    arr(0) = 1
    arr(1) = 2
    arr(2) = 3
    
    Dim varArr() As Variant
    varArr = Arrays.toVariantArray(arr)
    
    equals varArr(0), 1, "Turning a typed array to a variant array"
    equals varArr(1), 2, "Turning a typed array to a variant array"
    equals varArr(2), 3, "Turning a typed array to a variant array"
    
    
    arr = Arrays.emptyIntegerArray
    varArr = Arrays.toVariantArray(arr)
    equals Arrays.elems(varArr), 0, "Turning an empty typed array to a variant array"
    
    gStop
End Sub
