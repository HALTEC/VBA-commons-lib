VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 0 ' Default


Private Declare Function WideCharToMultiByte Lib "kernel32.dll" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long

Private Declare Function MultiByteToWideChar Lib "kernel32.dll" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long) As Long

Private Declare Function GetLastError Lib "kernel32.dll" () As Long

' For how to use the *W functions see: http://blog.nkadesign.com/2013/vba-unicode-strings-and-the-windows-api/
' In short: Declare as LongPtr, not String, pass String using StrPtr(), because VBA always converts its nice unicode strings
' to ANSI when it sees String in a Declare statement.
Private Declare Function PathIsRelativeW Lib "shlwapi.dll" ( _
    ByVal lpszPath As LongPtr) As Boolean

Private Declare Function GetTempPathW Lib "kernel32.dll" ( _
    ByVal nBufferLength As Long, _
    ByVal lpBuffer As LongPtr) As Long


Public Function pathIsRelative(ByVal path As String) As Boolean
    pathIsRelative = PathIsRelativeW(StrPtr(path))
End Function


Public Function pathJoin(ParamArray fragments() As Variant) As String
    Dim i As Integer
    Dim resultPath As String, t
    resultPath = ""
    For i = LBound(fragments) To UBound(fragments)
        If TypeName(fragments(i)) <> "String" Then
            Err.Raise E_INVALIDINPUT
        End If
        
        If fragments(i) <> "" Then
            If resultPath = "" Then
                resultPath = fragments(i)
            ElseIf Stringx.endsWith(resultPath, "\") And Stringx.startsWith(fragments(i), "\") Then
                resultPath = resultPath & Stringx.substr(fragments(i), 1)
            ElseIf Not Stringx.endsWith(resultPath, "\") And Not Stringx.startsWith(fragments(i), "\") Then
                resultPath = resultPath & "\" & fragments(i)
            Else
                resultPath = resultPath & fragments(i)
            End If
        End If
    Next
    
    pathJoin = resultPath
End Function


Public Function pathTempPath() As String
    ' According to https://msdn.microsoft.com/en-us/library/windows/desktop/aa364992(v=vs.85).aspx
    Dim bufferLength As Long: bufferLength = 261
    
    Dim buffer As String
    buffer = String$(bufferLength, "_")
    
    Dim pathLength As Long: pathLength = GetTempPathW(bufferLength, StrPtr(buffer))
    
    If pathLength = 0 Then
        Err.Raise E_INTERNALERROR
    Else
        If pathLength > bufferLength Then
            Err.Raise E_INTERNALERROR, description:="Path is longer than expected."
        End If
        
        buffer = Left$(buffer, pathLength)
    End If
    
    pathTempPath = buffer
End Function


Function dirExists(ByVal path As String) As Boolean
    On Error Resume Next
    dirExists = ((GetAttr(path) And vbDirectory) = vbDirectory)
    Err.clear
End Function


Function fileExists(ByVal path As String) As Boolean
    On Error Resume Next
    fileExists = ((GetAttr(path) And vbDirectory) <> vbDirectory)
    Err.clear
End Function


Public Function slurp(filename As String, _
        Optional enc As String = "utf-8", _
        Optional bin As Boolean = False) As Variant
    
    If Not fileExists(filename) Then
        Err.Raise E_FILENOTFOUND
    End If
    
    Dim fileNo As Integer: fileNo = FreeFile
    
    Open filename For Binary Access Read Lock Write As #fileNo
    
    If LOF(fileNo) = 0 Then
        If bin Then
            slurp = Arrays.emptyByteArray
        Else
            slurp = ""
        End If
        
        Exit Function
    End If
    
    Dim buffer() As Byte
    ReDim buffer(LOF(fileNo) - 1)
    
    Get #fileNo, , buffer
    
    Close #fileNo
    
    If bin Then
        slurp = buffer
    Else
        Dim text As String
        slurp = decode(buffer, enc)
    End If
End Function


Public Function spurt(filename As String, content As Variant, _
        Optional enc As String = "utf-8", _
        Optional append As Boolean = False, _
        Optional createOnly As Boolean = False) As String
    
    Dim fileNo As Integer: fileNo = FreeFile
    
    If fileExists(filename) Then
        If createOnly Then
            Err.Raise E_FILEEXISTS
        End If
        
        If Not append Then
            ' Trim file to size 0
            Open filename For Output As #fileNo
            Close #fileNo
        End If
    End If
    
    Open filename For Binary Access Write Lock Read Write As #fileNo
    
    Dim buffer() As Byte
    If varType(content) = (vbByte Or vbArray) Then
        buffer = content
    ElseIf varType(content) = vbString Then
        buffer = encode(CStr(content), enc)
    Else
        Err.Raise E_TYPEMISMATCH
    End If
    
    If append Then
        ' Put takes positions 1-based
        Put #fileNo, LOF(fileNo) + 1, buffer
    Else
        Put #fileNo, , buffer
    End If
    
    Close #fileNo
End Function


Public Function decode(ByRef inBuffer() As Byte, enc As String) As String
    If Arrays.elems(inBuffer) = 0 Then
        decode = ""
        Exit Function
    End If
    
    ' https://msdn.microsoft.com/en-us/library/windows/desktop/dd374130(v=vs.85).aspx
    Dim codepageNo As Long: codepageNo = encToCPNo(enc)
    
    
    ' Retrieve the required buffer size.
    Dim outStringSize As Long: outStringSize = MultiByteToWideChar(codepageNo, 0, VarPtr(inBuffer(LBound(inBuffer))), Arrays.elems(inBuffer), 0, 0)
    
    If outStringSize = 0 Then
        If GetLastError() = 1113 Then ' 1113 = ERROR_NO_UNICODE_TRANSLATION
            ' Thrown when invalid input was found.
            Err.Raise E_INVALIDINPUT
        Else
            Err.Raise E_INTERNALERROR
        End If
    End If
    
    Dim outString As String: outString = String$(outStringSize, 0)
    
    If MultiByteToWideChar(codepageNo, 0, VarPtr(inBuffer(LBound(inBuffer))), Arrays.elems(inBuffer), StrPtr(outString), outStringSize) = 0 Then
        Err.Raise E_INTERNALERROR
    End If
    
    decode = outString
End Function


Public Function encode(ByRef inString As String, enc As String) As Byte()
    If Len(inString) = 0 Then
        ' TODO: Test this!
        encode = Arrays.emptyByteArray()
        Exit Function
    End If
    
    ' https://msdn.microsoft.com/en-us/library/windows/desktop/dd374130(v=vs.85).aspx
    Dim codepageNo As Long: codepageNo = encToCPNo(enc)
    
    ' Retrieve the required buffer size.
    Dim bufferSize As Long: bufferSize = WideCharToMultiByte(codepageNo, 0, StrPtr(inString), Len(inString), 0, 0, 0, 0)
    
    If bufferSize = 0 Then
        If GetLastError() = 1113 Then ' 1113 = ERROR_NO_UNICODE_TRANSLATION
            ' Thrown when invalid input was found.
            Err.Raise E_INVALIDINPUT
        Else
            Err.Raise E_INTERNALERROR
        End If
    End If
    
    Dim outBuffer() As Byte
    ReDim outBuffer(0 To bufferSize - 1)
    
    If WideCharToMultiByte(codepageNo, 0, StrPtr(inString), Len(inString), VarPtr(outBuffer(0)), bufferSize, 0, 0) = 0 Then
        Err.Raise E_INTERNALERROR
    End If
    
    encode = outBuffer
End Function


Private Function encToCPNo(enc As String) As Long
    ' Uses the .NET Code Page Identifiers
    ' https://msdn.microsoft.com/en-us/library/windows/desktop/dd317756(v=vs.85).aspx
    Select Case enc
        Case "IBM037"
            encToCPNo = 37
        Case "IBM437"
            encToCPNo = 437
        Case "IBM500"
            encToCPNo = 500
        Case "ASMO-708"
            encToCPNo = 708
        Case "DOS-720"
            encToCPNo = 720
        Case "ibm737"
            encToCPNo = 737
        Case "ibm775"
            encToCPNo = 775
        Case "ibm850"
            encToCPNo = 850
        Case "ibm852"
            encToCPNo = 852
        Case "IBM855"
            encToCPNo = 855
        Case "ibm857"
            encToCPNo = 857
        Case "IBM00858"
            encToCPNo = 858
        Case "IBM860"
            encToCPNo = 860
        Case "ibm861"
            encToCPNo = 861
        Case "DOS-862"
            encToCPNo = 862
        Case "IBM863"
            encToCPNo = 863
        Case "IBM864"
            encToCPNo = 864
        Case "IBM865"
            encToCPNo = 865
        Case "cp866"
            encToCPNo = 866
        Case "ibm869"
            encToCPNo = 869
        Case "IBM870"
            encToCPNo = 870
        Case "windows-874"
            encToCPNo = 874
        Case "cp875"
            encToCPNo = 875
        Case "shift_jis"
            encToCPNo = 932
        Case "gb2312"
            encToCPNo = 936
        Case "ks_c_5601-1987"
            encToCPNo = 949
        Case "big5"
            encToCPNo = 950
        Case "IBM1026"
            encToCPNo = 1026
        Case "IBM01047"
            encToCPNo = 1047
        Case "IBM01140"
            encToCPNo = 1140
        Case "IBM01141"
            encToCPNo = 1141
        Case "IBM01142"
            encToCPNo = 1142
        Case "IBM01143"
            encToCPNo = 1143
        Case "IBM01144"
            encToCPNo = 1144
        Case "IBM01145"
            encToCPNo = 1145
        Case "IBM01146"
            encToCPNo = 1146
        Case "IBM01147"
            encToCPNo = 1147
        Case "IBM01148"
            encToCPNo = 1148
        Case "IBM01149"
            encToCPNo = 1149
        Case "utf-16"
            encToCPNo = 1200
        Case "unicodeFFFE"
            encToCPNo = 1201
        Case "windows-1250"
            encToCPNo = 1250
        Case "windows-1251"
            encToCPNo = 1251
        Case "windows-1252"
            encToCPNo = 1252
        Case "windows-1253"
            encToCPNo = 1253
        Case "windows-1254"
            encToCPNo = 1254
        Case "windows-1255"
            encToCPNo = 1255
        Case "windows-1256"
            encToCPNo = 1256
        Case "windows-1257"
            encToCPNo = 1257
        Case "windows-1258"
            encToCPNo = 1258
        Case "Johab"
            encToCPNo = 1361
        Case "macintosh"
            encToCPNo = 10000
        Case "x-mac-japanese"
            encToCPNo = 10001
        Case "x-mac-chinesetrad"
            encToCPNo = 10002
        Case "x-mac-korean"
            encToCPNo = 10003
        Case "x-mac-arabic"
            encToCPNo = 10004
        Case "x-mac-hebrew"
            encToCPNo = 10005
        Case "x-mac-greek"
            encToCPNo = 10006
        Case "x-mac-cyrillic"
            encToCPNo = 10007
        Case "x-mac-chinesesimp"
            encToCPNo = 10008
        Case "x-mac-romanian"
            encToCPNo = 10010
        Case "x-mac-ukrainian"
            encToCPNo = 10017
        Case "x-mac-thai"
            encToCPNo = 10021
        Case "x-mac-ce"
            encToCPNo = 10029
        Case "x-mac-icelandic"
            encToCPNo = 10079
        Case "x-mac-turkish"
            encToCPNo = 10081
        Case "x-mac-croatian"
            encToCPNo = 10082
        Case "utf-32"
            encToCPNo = 12000
        Case "utf-32BE"
            encToCPNo = 12001
        Case "x-Chinese_CNS"
            encToCPNo = 20000
        Case "x-cp20001"
            encToCPNo = 20001
        Case "x_Chinese-Eten"
            encToCPNo = 20002
        Case "x-cp20003"
            encToCPNo = 20003
        Case "x-cp20004"
            encToCPNo = 20004
        Case "x-cp20005"
            encToCPNo = 20005
        Case "x-IA5"
            encToCPNo = 20105
        Case "x-IA5-German"
            encToCPNo = 20106
        Case "x-IA5-Swedish"
            encToCPNo = 20107
        Case "x-IA5-Norwegian"
            encToCPNo = 20108
        Case "us-ascii"
            encToCPNo = 20127
        Case "x-cp20261"
            encToCPNo = 20261
        Case "x-cp20269"
            encToCPNo = 20269
        Case "IBM273"
            encToCPNo = 20273
        Case "IBM277"
            encToCPNo = 20277
        Case "IBM278"
            encToCPNo = 20278
        Case "IBM280"
            encToCPNo = 20280
        Case "IBM284"
            encToCPNo = 20284
        Case "IBM285"
            encToCPNo = 20285
        Case "IBM290"
            encToCPNo = 20290
        Case "IBM297"
            encToCPNo = 20297
        Case "IBM420"
            encToCPNo = 20420
        Case "IBM423"
            encToCPNo = 20423
        Case "IBM424"
            encToCPNo = 20424
        Case "x-EBCDIC-KoreanExtended"
            encToCPNo = 20833
        Case "IBM-Thai"
            encToCPNo = 20838
        Case "koi8-r"
            encToCPNo = 20866
        Case "IBM871"
            encToCPNo = 20871
        Case "IBM880"
            encToCPNo = 20880
        Case "IBM905"
            encToCPNo = 20905
        Case "IBM00924"
            encToCPNo = 20924
        Case "EUC-JP"
            encToCPNo = 20932
        Case "x-cp20936"
            encToCPNo = 20936
        Case "x-cp20949"
            encToCPNo = 20949
        Case "cp1025"
            encToCPNo = 21025
        Case "koi8-u"
            encToCPNo = 21866
        Case "iso-8859-1"
            encToCPNo = 28591
        Case "iso-8859-2"
            encToCPNo = 28592
        Case "iso-8859-3"
            encToCPNo = 28593
        Case "iso-8859-4"
            encToCPNo = 28594
        Case "iso-8859-5"
            encToCPNo = 28595
        Case "iso-8859-6"
            encToCPNo = 28596
        Case "iso-8859-7"
            encToCPNo = 28597
        Case "iso-8859-8"
            encToCPNo = 28598
        Case "iso-8859-9"
            encToCPNo = 28599
        Case "iso-8859-13"
            encToCPNo = 28603
        Case "iso-8859-15"
            encToCPNo = 28605
        Case "x-Europa"
            encToCPNo = 29001
        Case "iso-8859-8-i"
            encToCPNo = 38598
        Case "iso-2022-jp"
            encToCPNo = 50220
        Case "csISO2022JP"
            encToCPNo = 50221
        Case "iso-2022-jp"
            encToCPNo = 50222
        Case "iso-2022-kr"
            encToCPNo = 50225
        Case "x-cp50227"
            encToCPNo = 50227
        Case "euc-jp"
            encToCPNo = 51932
        Case "EUC-CN"
            encToCPNo = 51936
        Case "euc-kr"
            encToCPNo = 51949
        Case "hz-gb-2312"
            encToCPNo = 52936
        Case "GB18030"
            encToCPNo = 54936
        Case "x-iscii-de"
            encToCPNo = 57002
        Case "x-iscii-be"
            encToCPNo = 57003
        Case "x-iscii-ta"
            encToCPNo = 57004
        Case "x-iscii-te"
            encToCPNo = 57005
        Case "x-iscii-as"
            encToCPNo = 57006
        Case "x-iscii-or"
            encToCPNo = 57007
        Case "x-iscii-ka"
            encToCPNo = 57008
        Case "x-iscii-ma"
            encToCPNo = 57009
        Case "x-iscii-gu"
            encToCPNo = 57010
        Case "x-iscii-pa"
            encToCPNo = 57011
        Case "utf-7"
            encToCPNo = 65000
        Case "utf-8"
            encToCPNo = 65001
        Case Default
            Err.Raise E_UNKNOWNENCODING
    End Select
End Function
