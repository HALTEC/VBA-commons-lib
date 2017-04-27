Attribute VB_Name = "Test_Variants"
Option Explicit

Private Sub test_all()
    gStart "Variants"
    
    test_verifyArray
    test_fuzzyEquals
    
    gStop
End Sub

Private Sub test_verifyArray()
    gStart "verifyArray"
    
    On Error Resume Next
    Variants.verifyArray "Ducky"
    checkError E_ARGUMENTOUTOFRANGE
    On Error GoTo 0
    
    On Error Resume Next
    Variants.verifyArray CVar("Ducky")
    checkError E_ARGUMENTOUTOFRANGE
    On Error GoTo 0
    
    Dim arr(0 To 2) As String
    arr(0) = "Hey"
    arr(1) = "Boo"
    On Error Resume Next
    Variants.verifyArray arr
    checkNoError
    On Error GoTo 0
    
    gStop
End Sub


Sub test_fuzzyEquals()
    gStart "fuzzy equals"

    On Error Resume Next
    Variants.fuzzyEquals Variants, Variants
    checkError E_ARGUMENTOUTOFRANGE, "Throws on things that are not equatable."
    On Error GoTo 0

    ok Variants.fuzzyEquals(5, 5)
    ok Not Variants.fuzzyEquals(5, 4)
    ok Not Variants.fuzzyEquals("5", 5)
    ok Variants.fuzzyEquals("Hi", "Hi")
    
    gStop
End Sub
