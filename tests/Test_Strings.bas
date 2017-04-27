Attribute VB_Name = "Test_Strings"
Option Explicit
Option Base 0 ' Default

Private Sub test_all()
    gStart "Strings"
    
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
    
    equals Strings.startsWith("asdf", "as"), True, "normal"
    equals Strings.startsWith("asdf", ""), True, "empty prefix"
    equals Strings.startsWith("asdf", "asdfgh"), False, "longer prefix than text"
    
    gStop
End Sub

Private Sub test_endsWith()
    gStart "endsWith"
    
    equals Strings.endsWith("asdf", "df"), True, "normal"
    equals Strings.endsWith("asdf", ""), True, "empty postfix"
    equals Strings.endsWith("asdf", "xzasdf"), False, "longer postfix than text"
    
    gStop
End Sub

Private Sub test_trimChar()
    gStart "trimChar"
    
    equals Strings.trimChar(" as df "), "as df", "Normales trim geht"
    
    gStop
End Sub

Private Sub test_trimCharStart()
    gStart "trimCharStart"
    
    equals Strings.trimCharStart(" asdf "), "asdf "
    equals Strings.trimCharStart(" " & vbVerticalTab & "asdf"), "asdf"
    equals Strings.trimCharStart("as df"), "as df"
    equals Strings.trimCharStart(",as,df", ","), "as,df"
    equals Strings.trimCharStart(",as,df", ",;"), "as,df"
    equals Strings.trimCharStart(";,as,df", ",;"), "as,df"
    
    gStop
End Sub

Private Sub test_trimCharEnd()
    gStart "trimCharEnd"
    
    equals Strings.trimCharEnd(" asdf "), " asdf"
    equals Strings.trimCharEnd("asdf" & vbVerticalTab & " "), "asdf"
    equals Strings.trimCharEnd("as df"), "as df"
    equals Strings.trimCharEnd("as,df,", ","), "as,df"
    equals Strings.trimCharEnd("as,df,", ",;"), "as,df"
    equals Strings.trimCharEnd("as,df;,", ",;"), "as,df"
    
    gStop
End Sub

Private Sub test_substr()
    gStart "substr"
    
    equals Strings.substr("abcd", 0), "abcd"
    equals Strings.substr("abcd", 1), "bcd"
    equals Strings.substr("abcd", 0, 1), "a"
    equals Strings.substr("abcd", 0, 4), "abcd"
    equals Strings.substr("abcd", 0, -1), "abc"
    equals Strings.substr("abcd", -3, -1), "bc"
    equals Strings.substr("abcd", -3, 1), "b"
    equals Strings.substr("abcd", 2, 0), "", "Zero length substr returns empty string"
    
    On Error Resume Next
    Strings.substr "abcd", 3, -2
    checkError E_INDEXOUTOFRANGE, "Length that results in stop index < start index throws"
    On Error GoTo 0
    
    gStop
End Sub

Private Sub test_split()
    gStart "split"
    
    equals Strings.split("asdf, qwer, 1", ", "), List_createLT("asdf", "qwer", "1")
    equals Strings.split("", ", "), List_createLT(""), "Empty string"
    equals Strings.split(", ", ", "), List_createLT("", ""), "Single separator"
    equals Strings.split(", asdf", ", "), List_createLT("", "asdf"), "Leading empty element"
    equals Strings.split("abc", ""), List_createLT("a", "b", "c"), "Empty separator splits every char"
    
    gStop
End Sub

Private Sub test_join()
    gStart "join"
    
    equals Strings.join(List_create("asdf", "qwer", "1"), ", "), "asdf, qwer, 1"
    equals Strings.join(List_create, ", "), "", "Empty list"
    equals Strings.join(List_create(""), ", "), "", "One empty string"
    equals Strings.join(List_create("", ""), ", "), ", ", "Two empty strings"
    equals Strings.join(List_create("", "asdf"), ", "), ", asdf", "Leading empty element results in a leading separator"
    
    gStop
End Sub

Private Sub test_toChars()
    gStart "toChars"
    
    equals Strings.toChars("asdf"), List_create("a", "s", "d", "f")
    equals Strings.toChars(""), List_create
    equals Strings.toChars("I! AM! Sparta."), List_create("I", "!", " ", "A", "M", "!", " ", "S", "p", "a", "r", "t", "a", ".")
    
    gStop
End Sub

Private Sub test_repeat()
    gStart "repeat"
    
    equals Strings.repeat("a", 0), ""
    equals Strings.repeat("a", 1), "a"
    equals Strings.repeat("a", 2), "aa"
    equals Strings.repeat("ab", 2), "abab", "Works with text with more than one char"
    equals Strings.repeat("", 5), "", "Empty text works"
    
    On Error Resume Next
    Strings.repeat "a", -1
    checkError E_INDEXOUTOFRANGE, "Negative repetition throws"
    On Error GoTo 0
    
    gStop
End Sub
