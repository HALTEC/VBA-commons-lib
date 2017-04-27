Attribute VB_Name = "Test_Math"
Option Explicit

Private Sub test_all()
    gStart "Math"
    
    test_min
    test_max
    test_cmp
    
    gStop
End Sub

Private Sub test_min()
    gStart "min"
    
    equals Math.min(1, 2), 1
    equals Math.min(2, 1), 1
    
    equals Math.min("a", "b"), "a"
    equals Math.min("b", "a"), "a"
    
    equals Math.min(2, 1, 3, 4), 1
    
    gStop
End Sub

Private Sub test_max()
    gStart "max"
    
    equals Math.max(1, 2), 2
    equals Math.max(2, 1), 2
    
    equals Math.max("a", "b"), "b"
    equals Math.max("b", "a"), "b"
    
    equals Math.max(2, 1, 3, 4), 4
    
    gStop
End Sub

Private Sub test_cmp()
    gStart "cmp"
    
    equals Math.cmp(1, 2), -1
    equals Math.cmp(2, 1), 1
    equals Math.cmp(1, 1), 0
    
    equals Math.cmp("a", "b"), -1
    equals Math.cmp("b", "a"), 1
    equals Math.cmp("a", "a"), 0
    
    gStop
End Sub

