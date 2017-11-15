Attribute VB_Name = "Test_Stringx"
Option Explicit
Option Base 0 ' Default

Private Sub test_all()
    gStart "Stringx"
    
    test_startsWith
    test_endsWith
    test_trimChar
    test_trimCharStart
    test_trimCharEnd
    test_substr
    test_split
    test_join
    test_toChars
    
    gStop
End Sub


Private Sub test_startsWith()
    gStart "startsWith"
    
    equals Stringx.startsWith("asdf", "as"), True, "normal"
    equals Stringx.startsWith("asdf", ""), True, "empty prefix"
    equals Stringx.startsWith("asdf", "asdfgh"), False, "longer prefix than text"
    
    gStop
End Sub

Private Sub test_endsWith()
    gStart "endsWith"
    
    equals Stringx.endsWith("asdf", "df"), True, "normal"
    equals Stringx.endsWith("asdf", ""), True, "empty postfix"
    equals Stringx.endsWith("asdf", "xzasdf"), False, "longer postfix than text"
    
    gStop
End Sub

Private Sub test_trimChar()
    gStart "trimChar"
    
    equals Stringx.trimChar(" as df "), "as df", "Normales trim geht"
    equals Stringx.trimChar(""), ""
    
    gStop
End Sub

Private Sub test_trimCharStart()
    gStart "trimCharStart"
    
    equals Stringx.trimCharStart(" asdf "), "asdf "
    equals Stringx.trimCharStart(" " & vbVerticalTab & "asdf"), "asdf"
    equals Stringx.trimCharStart("as df"), "as df"
    equals Stringx.trimCharStart(",as,df", ","), "as,df"
    equals Stringx.trimCharStart(",as,df", ",;"), "as,df"
    equals Stringx.trimCharStart(";,as,df", ",;"), "as,df"
    equals Stringx.trimCharStart(""), ""
    
    gStop
End Sub

Private Sub test_trimCharEnd()
    gStart "trimCharEnd"
    
    equals Stringx.trimCharEnd(" asdf "), " asdf"
    equals Stringx.trimCharEnd("asdf" & vbVerticalTab & " "), "asdf"
    equals Stringx.trimCharEnd("as df"), "as df"
    equals Stringx.trimCharEnd("as,df,", ","), "as,df"
    equals Stringx.trimCharEnd("as,df,", ",;"), "as,df"
    equals Stringx.trimCharEnd("as,df;,", ",;"), "as,df"
    equals Stringx.trimCharEnd(""), ""
    
    gStop
End Sub

Private Sub test_substr()
    gStart "substr"
    
    equals Stringx.substr("abcd", 0), "abcd"
    equals Stringx.substr("abcd", 1), "bcd"
    equals Stringx.substr("abcd", 0, 1), "a"
    equals Stringx.substr("abcd", 0, 4), "abcd"
    equals Stringx.substr("abcd", 0, -1), "abc"
    equals Stringx.substr("abcd", -3, -1), "bc"
    equals Stringx.substr("abcd", -3, 1), "b"
    equals Stringx.substr("abcd", 2, 0), "", "Zero length substr returns empty string"
    
    On Error Resume Next
    Stringx.substr "abcd", 3, -2
    checkError E_INDEXOUTOFRANGE, "Length that results in stop index < start index throws"
    On Error GoTo 0
    
    gStop
End Sub

Private Sub test_split()
    gStart "split"
    
    equals Stringx.split("asdf, qwer, 1", ", "), List_createLT("asdf", "qwer", "1")
    equals Stringx.split("", ", "), List_createLT(""), "Empty string"
    equals Stringx.split(", ", ", "), List_createLT("", ""), "Single separator"
    equals Stringx.split(", asdf", ", "), List_createLT("", "asdf"), "Leading empty element"
    equals Stringx.split("abc", ""), List_createLT("a", "b", "c"), "Empty separator splits every char"
    
    gStop
End Sub

Private Sub test_join()
    gStart "join"
    
    equals Stringx.join(List_create("asdf", "qwer", "1"), ", "), "asdf, qwer, 1"
    equals Stringx.join(List_create, ", "), "", "Empty list"
    equals Stringx.join(List_create(""), ", "), "", "One empty string"
    equals Stringx.join(List_create("", ""), ", "), ", ", "Two empty strings"
    equals Stringx.join(List_create("", "asdf"), ", "), ", asdf", "Leading empty element results in a leading separator"
    
    gStop
End Sub

Private Sub test_toChars()
    gStart "toChars"
    
    equals Stringx.toChars("asdf"), List_createLT("a", "s", "d", "f")
    equals Stringx.toChars(""), List_createT("String")
    equals Stringx.toChars("I! AM! Sparta."), List_createLT("I", "!", " ", "A", "M", "!", " ", "S", "p", "a", "r", "t", "a", ".")
    
    gStop
End Sub

Private Sub test_repeat()
    gStart "repeat"
    
    equals Stringx.repeat("a", 0), ""
    equals Stringx.repeat("a", 1), "a"
    equals Stringx.repeat("a", 2), "aa"
    equals Stringx.repeat("ab", 2), "abab", "Works with text with more than one char"
    equals Stringx.repeat("", 5), "", "Empty text works"
    
    On Error Resume Next
    Stringx.repeat "a", -1
    checkError E_INDEXOUTOFRANGE, "Negative repetition throws"
    On Error GoTo 0
    
    gStop
End Sub
