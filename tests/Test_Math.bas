Attribute VB_Name = "Test_Math"
Option Explicit

Private Sub test_all()
    gStart "Math"
    
    test_min
    test_max
    test_cmp
    test_ceiling
    test_floor
    
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

Private Sub test_ceiling()
    gStart "ceiling"
    
    equals Math.ceiling(1), 1
    equals Math.ceiling(0), 0
    equals Math.ceiling(1.5), 2
    equals Math.ceiling(12345.123), 12346
    
    equals Math.ceiling(-1), -1
    equals Math.ceiling(-1.5), -1
    equals Math.ceiling(-12345.123), -12345
    
    gStop
End Sub

Private Sub test_floor()
    gStart "floor"
    
    equals Math.floor(1), 1
    equals Math.floor(0), 0
    equals Math.floor(1.5), 1
    equals Math.floor(12345.123), 12345
    
    equals Math.floor(-1), -1
    equals Math.floor(-1.5), -2
    equals Math.floor(-12345.123), -12346
    
    gStop
End Sub
