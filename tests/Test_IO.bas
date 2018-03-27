Attribute VB_Name = "Test_IO"
Option Explicit
Option Base 0 ' Default

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" ( _
    ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Const testFolder = "C:\data\repos\VBA-commons-lib\tests\IO\"

Private Sub test_all()
    gStart "IO"
    
    test_pathIsRelative
    test_pathJoin
    test_encode
    test_decode
    test_slurp
    test_spurt
    
    gStop
End Sub

Private Sub test_pathIsRelative()
    gStart "pathIsRelative"
    
    equals IO.pathIsRelative("one"), True, "simple relative path"
    equals IO.pathIsRelative("C:\one"), False, "simple absolute path"
    equals IO.pathIsRelative("\one"), False, "leading backslash is absolute"
    equals IO.pathIsRelative("C:\one\two\..\three\"), False, "complex absolute path"
    equals IO.pathIsRelative("one\two\..\three\"), True, "complex relative path"
    
    gStop
End Sub

Private Sub test_pathJoin()
    gStart "pathJoin"
    
    equals IO.pathJoin("one", "two"), "one\two", "no separators in between"
    equals IO.pathJoin("one\", "two"), "one\two", "separator after path 1"
    equals IO.pathJoin("one", "\two"), "one\two", "separator before path 2"
    equals IO.pathJoin("one\", "\two"), "one\two", "separator after path 1 and before path 2"
    
    equals IO.pathJoin("", "two"), "two", "empty path 1"
    equals IO.pathJoin("", "C:\two"), "C:\two", "empty path 1 and absolute path 2"
    equals IO.pathJoin("C:\one", ""), "C:\one", "empty path 1 and absolute path 2"
    
    equals IO.pathJoin("C:\one", "two", "three", "four"), "C:\one\two\three\four", "multiple paths"
    
    equals IO.pathJoin("\one", "two"), "\one\two", "leading backslash"
    
    equals IO.pathJoin("\\one\\", "\two\\"), "\\one\\two\\", "excess backslashes"
    
    gStop
End Sub

Private Sub test_encode()
    gStart "encode"
    
    Dim b() As Byte
    
    b = x("a", "s", "d", "f")
    equals IO.encode("asdf", "utf-8"), b, "utf-8 basic string"
    
    b = x()
    equals IO.encode("", "utf-8"), b, "utf-8 empty string"
    
    ' Unicode Character 'GREEK SMALL LETTER ALPHA' (U+03B1) = 945
    b = x(&HCE, &HB1)
    equals IO.encode(ChrW(945), "utf-8"), b, "utf-8 non ascii char"
    
    b = x(&HCE, &HB1, &HCE, &HB1)
    equals IO.encode(ChrW(945) & ChrW(945), "utf-8"), b, "utf-8 multiple non ascii chars"
    
    b = x("a", vbCr, vbLf, "s")
    equals IO.encode("a" & vbNewLine & "s", "utf-8"), b, "utf-8 newline"
    
    b = x("a", "s", "d", "f")
    equals IO.encode("asdf", "us-ascii"), b, "us-ascii encoding"
    
    b = x("a", vbCr, vbLf, "s")
    equals IO.encode("a" & vbNewLine & "s", "us-ascii"), b, "us-ascii newline"
    
    gStop
End Sub

Private Sub test_decode()
    gStart "decode"
    
    Dim b() As Byte
    
    b = x("a", "s", "d", "f")
    equals IO.decode(b, "utf-8"), "asdf", "utf-8 basic string"
    
    b = x()
    equals IO.decode(b, "utf-8"), "", "utf-8 empty string"
    
    ' Unicode Character 'GREEK SMALL LETTER ALPHA' (U+03B1) = 945
    b = x(&HCE, &HB1)
    equals IO.decode(b, "utf-8"), ChrW(945), "utf-8 non ascii char"
    
    b = x(&HCE, &HB1, &HCE, &HB1)
    equals IO.decode(b, "utf-8"), ChrW(945) & ChrW(945), "utf-8 multiple non ascii chars"
    
    b = x("a", vbCr, vbLf, "s")
    equals IO.decode(b, "utf-8"), "a" & vbNewLine & "s", "utf-8 newline"
    
    b = x("a", "s", "d", "f")
    equals IO.decode(b, "us-ascii"), "asdf", "us-ascii encoding"
    
    b = x("a", vbCr, vbLf, "s")
    equals IO.decode(b, "us-ascii"), "a" & vbNewLine & "s", "us-ascii newline"
    
    gStop
End Sub

Private Sub test_slurp()
    gStart "slurp"
    
    equals IO.slurp(testFolder & "a.txt"), "", "empty file"
    
    equals IO.slurp(testFolder & "b.txt"), "asdf", "ascii file"
    
    equals IO.slurp(testFolder & "c.txt"), "one" & vbCrLf & "two", "ascii file with windows newline"
    
    equals IO.slurp(testFolder & "d.txt"), "one" & vbLf & "two", "ascii file with unix newline"
    
    equals IO.slurp(testFolder & "e.txt"), "asdf" & vbLf, "ascii file with trailing newline"
    
    equals IO.slurp(testFolder & "f.txt"), "Größenwahn", "utf-8 non-ascii file"
    
    equals IO.slurp(testFolder & "g.txt", enc:="iso-8859-1"), "Größenwahn", "iso-8859-1 non-ascii file"
    
    equals IO.slurp(testFolder & "a.txt", bin:=True), Arrays.emptyByteArray, "binary empty file"
    
    equals IO.slurp(testFolder & "b.txt", bin:=True), x("a", "s", "d", "f"), "binary ascii file"
    
    gStop
End Sub

Private Sub test_spurt()
    gStart "spurt"
    
    On Error Resume Next
    Kill testFolder & "out_*.txt"
    On Error GoTo 0
    
    IO.spurt testFolder & "out_a.txt", ""
    ok areFilesEqual(testFolder & "a.txt", testFolder & "out_a.txt"), "empty file"
    
    IO.spurt testFolder & "out_b.txt", "asdf"
    ok areFilesEqual(testFolder & "b.txt", testFolder & "out_b.txt"), "ascii file"
    
    IO.spurt testFolder & "out_c.txt", "one" & vbCrLf & "two"
    ok areFilesEqual(testFolder & "c.txt", testFolder & "out_c.txt"), "ascii file with windows newline"
    
    IO.spurt testFolder & "out_d.txt", "one" & vbLf & "two"
    ok areFilesEqual(testFolder & "d.txt", testFolder & "out_d.txt"), "ascii file with unix newline"
    
    IO.spurt testFolder & "out_e.txt", "asdf" & vbLf
    ok areFilesEqual(testFolder & "e.txt", testFolder & "out_e.txt"), "ascii file with trailing newline"
    
    IO.spurt testFolder & "out_f.txt", "Größenwahn"
    ok areFilesEqual(testFolder & "f.txt", testFolder & "out_f.txt"), "utf-8 non-ascii file"
    
    IO.spurt testFolder & "out_g.txt", "Größenwahn", enc:="iso-8859-1"
    ok areFilesEqual(testFolder & "g.txt", testFolder & "out_g.txt"), "iso-8859-1 non-ascii file"
    
    IO.spurt testFolder & "out_bin_a.txt", Arrays.emptyByteArray
    ok areFilesEqual(testFolder & "a.txt", testFolder & "out_bin_a.txt"), "binary empty file"
    
    IO.spurt testFolder & "out_bin_b.txt", x("a", "s", "d", "f")
    ok areFilesEqual(testFolder & "b.txt", testFolder & "out_bin_b.txt"), "binary ascii file"
    
    IO.spurt testFolder & "out_h.txt", "One two three."
    IO.spurt testFolder & "out_h.txt", vbLf & "Four five seven.", append:=True
    ok areFilesEqual(testFolder & "h.txt", testFolder & "out_h.txt"), "appending"
    
    IO.spurt testFolder & "out_i.txt", "This text will be overwritten."
    IO.spurt testFolder & "out_i.txt", "This is the new text."
    ok areFilesEqual(testFolder & "i.txt", testFolder & "out_i.txt"), "overwriting"
    
    IO.spurt testFolder & "out_j.txt", "Foobar", createOnly:=True
    ok areFilesEqual(testFolder & "j.txt", testFolder & "out_j.txt"), "createOnly w/o existing file"
    
    IO.spurt testFolder & "out_k.txt", "Stuff"
    On Error Resume Next
    IO.spurt testFolder & "out_k.txt", "New stuff", createOnly:=True
    checkError E_FILEEXISTS, "createOnly with existing file fails"
    On Error GoTo 0
    
    On Error Resume Next
    Kill testFolder & "out_*.txt"
    On Error GoTo 0
    
    gStop
End Sub

Private Function ShellX(ByVal PathName As String, _
        Optional ByVal WindowStyle As Long = vbMinimizedFocus, _
        Optional ByVal Events As Boolean = True _
        ) As Long
    Const STILL_ACTIVE = &H103&
    Const PROCESS_QUERY_INFORMATION = &H400&
    Dim procId As Long: procId = Shell(PathName, WindowStyle)
    Dim procHnd As Long: procHnd = OpenProcess(PROCESS_QUERY_INFORMATION, True, procId)
    
    Do
        If Events Then DoEvents
        GetExitCodeProcess procHnd, ShellX
    Loop While ShellX = STILL_ACTIVE
    
    CloseHandle procHnd
End Function

Public Function areFilesEqual(file1 As String, file2 As String) As Boolean
    areFilesEqual = ShellX("fc.exe /B " & file1 & " " & file2, Events:=False) = 0
End Function

Private Function x(ParamArray values() As Variant) As Byte()
    Dim arr() As Byte
    
    If UBound(values) - LBound(values) + 1 = 0 Then
        x = Arrays.emptyByteArray
        Exit Function
    End If
    
    ReDim arr(UBound(values) - LBound(values))
    
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If varType(values(i)) = vbString Then
            ' For some reason this does work on unicode systems for ascii range chars.
            arr(i) = AscB(values(i))
        ElseIf IsNumeric(values(i)) Then
            If values(i) >= 0 And values(i) < 256 Then
                arr(i) = CByte(values(i))
            Else
                Err.Raise E_ARGUMENTOUTOFRANGE
            End If
        Else
            Err.Raise E_ARGUMENTOUTOFRANGE
        End If
    Next

    x = arr
End Function
