VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stringx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private PADDING_CHAR As String
Private escapes As New Collection

Option Explicit
Option Base 0 ' Default

Public Function format(format_string As String, ParamArray values()) As String
'VB6 implementation of .net String.Format(), slightly customized.

        Dim return_value As String
        Dim values_count As Integer
        
        'some error-handling constants:
        Const ERR_FORMAT_EXCEPTION As Long = vbObjectError Or 9001
        Const ERR_ARGUMENT_NULL_EXCEPTION As Long = vbObjectError Or 9002
        Const ERR_SOURCE As String = "StringFormat"
        Const ERR_MSG_INVALID_FORMAT_STRING As String = "Invalid format string."
        Const ERR_MSG_FORMAT_EXCEPTION As String = "The number indicating an argument to format is less than zero, or greater than or equal to the length of the args array."
        
        'use SPACE as default padding character
        If PADDING_CHAR = vbNullString Then PADDING_CHAR = Chr$(32)
        
        'figure out number of passed values:
        values_count = UBound(values) + 1
                    
        Dim regex As RegExp
        Dim matches As MatchCollection
        Dim thisMatch As match
        Dim thisString As String
        Dim thisFormat As String
        
        'validate string_format:
        Set regex = New RegExp
        regex.pattern = "{({{)*(\w+)(,-?\d+)?(:[^}]+)?}(}})*"
        regex.ignoreCase = True
        regex.Global = True
        Set matches = regex.Execute(format_string)
        
        'determine if values_count matches number of unique regex matches:
        Dim uniqueCount As Integer
        Dim tmpCSV As String
        For Each thisMatch In matches
            If Not Stringx.contains(tmpCSV, thisMatch.SubMatches(1), False) Then
                uniqueCount = uniqueCount + 1
                tmpCSV = tmpCSV & thisMatch.SubMatches(1) & ","
            End If
        Next
        
        'unique indices count must match values_count:
        If matches.count > 0 And uniqueCount <> values_count Then _
            Error.RaiseInvalidFormatError ERR_SOURCE, "Unique indices mismatch values count."
        
        If Stringx.contains(format_string, "\\") Then _
            format_string = Replace(format_string, "\\", Chr$(27))
    
        If matches.count = 0 And format_string <> vbNullString And UBound(values) = -1 Then
        'only format_string was specified: skip to checking escape sequences:
            return_value = format_string
            GoTo checkEscapes
        ElseIf UBound(values) = -1 And matches.count > 0 Then
            Error.RaiseArgumentEmptyError ERR_SOURCE, "Format specifier has no value."
        End If
        
        return_value = format_string
    
        'dissect format_string:
        
        Dim i As Integer, v As String, p As String 'i: iterator; v: value; p: placeholder
        Dim alignmentGroup As String, alignmentSpecifier As String
        Dim formattedValue As String, alignmentPadding As Integer
        
        'iterate regex matches (each match is a placeholder):
        For i = 0 To matches.count - 1
            
            'get the placeholder specified index:
            Set thisMatch = matches(i)
            p = thisMatch.SubMatches(1)
            
            'if specified index (0-based) > uniqueCount (1-based), something's wrong:
            If p > uniqueCount - 1 Then _
                Error.RaiseInvalidFormatError ERR_SOURCE, "Format specifier index out of bounds."
            v = values(p)
            
            'get the alignment specifier if it is specified:
            alignmentGroup = thisMatch.SubMatches(2)
            If alignmentGroup <> vbNullString Then _
                alignmentSpecifier = Right$(alignmentGroup, LenB(alignmentGroup) / 2 - 1)
            
            
            'get the format specifier if it is specified:
            thisString = thisMatch.value
            If Stringx.contains(thisString, ":") Then
                
                Dim formatGroup As String, precisionSpecifier As Integer
                Dim formatSpecifier As String, precisionString As String
                
                'get the string between ":" and "}":
                formatGroup = Mid$(thisString, InStr(1, thisString, ":") + 1, (LenB(thisString) / 2) - 2)
                formatGroup = Left$(formatGroup, LenB(formatGroup) / 2 - 1)
                
                precisionString = Right$(formatGroup, LenB(formatGroup) / 2 - 1)
                formatSpecifier = Mid$(thisString, InStr(1, thisString, ":") + 1, 1)
                                
                'applicable formatting depends on the type of the value (yes, GOTO!!):
                If TypeName(values(p)) = "Date" Then GoTo DateTimeFormatSpecifiers
                If v = vbNullString Then GoTo ApplyStringFormat
                
NumberFormatSpecifiers:
                If precisionString <> vbNullString And Not IsNumeric(precisionString) Then _
                    Err.Raise ERR_FORMAT_EXCEPTION, _
                        ERR_SOURCE, ERR_MSG_INVALID_FORMAT_STRING
                
                If precisionString = vbNullString Then precisionString = 0
                
                Select Case formatSpecifier
                
                    Case "C", "c" 'CURRENCY format, formats string as currency.
                    'Precision specifier determines number of decimal digits.
                    'This implementation ignores regional settings
                    '(hard-coded group separator, decimal separator and currency sign).
                    
                    precisionSpecifier = CInt(precisionString)
                    thisFormat = "#,##0.00$"
                    
                    If LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                        'if a non-zero precision is specified...
                        thisFormat = Replace$(thisFormat, ".00", "." & String$(precisionString, Chr$(48)))
                    ElseIf LenB(formatGroup) > 2 And precisionSpecifier = 0 Then
                        thisFormat = Replace$(thisFormat, ".00", vbNullString)
                    End If
                    
                    
                    Case "D", "d" 'DECIMAL format, formats string as integer number.
                    'Precision specifier determines number of digits in returned string.
                    
                    
                    precisionSpecifier = CInt(precisionString)
                    thisFormat = "0"
                    thisFormat = Right$(String$(precisionSpecifier, "0") & thisFormat, _
                        IIf(precisionSpecifier = 0, Len(thisFormat), precisionSpecifier))
                    
                    
                    Case "E", "e" 'EXPONENTIAL NOTATION format (aka "Scientific Notation")
                    'Precision specifier determines number of decimals in returned string.
                    'This implementation ignores regional settings'
                    '(hard-coded decimal separator).
                    
                    
                    precisionSpecifier = CInt(precisionString)
                    thisFormat = "0.00000#" & formatSpecifier & "-#" 'defaults to 6 decimals
                    
                    If LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                        'if a non-zero precision is specified...
                        thisFormat = "0." & String$(precisionSpecifier - 1, Chr$(48)) & "#" & formatSpecifier & "-#"
                    
                    ElseIf LenB(formatGroup) > 2 And precisionSpecifier = 0 Then
                        Err.Raise ERR_FORMAT_EXCEPTION, _
                            ERR_SOURCE, ERR_MSG_INVALID_FORMAT_STRING
                    End If
                    
                    
                    Case "F", "f" 'FIXED-POINT format
                    'Precision specifier determines number of decimals in returned string.
                    'This implementation ignores regional settings'
                    '(hard-coded decimal separator).
                    
                    precisionSpecifier = CInt(precisionString)
                    thisFormat = "0"
                    If LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                        'if a non-zero precision is specified...
                        thisFormat = (thisFormat & ".") & String$(precisionSpecifier, Chr$(48))
                    Else
                        'no precision specified - default to 2 decimals:
                        thisFormat = "0.00"
                    End If
                    
                    
                    Case "G", "g" 'GENERAL format (recursive)
                    'returns the shortest of either FIXED-POINT or SCIENTIFIC formats in case of a Double.
                    'returns DECIMAL format in case of a Integer or Long.
                    
                    Dim eNotation As String, ePower As Integer, specifier As String
                    precisionSpecifier = IIf(CInt(precisionString) > 0, CInt(precisionString), _
                        IIf(Stringx.contains(v, "."), Len(v) - InStr(1, v, "."), 0))
                    
                    'track character case of formatSpecifier:
                    specifier = IIf(formatSpecifier = "G", "D", "d")
                    
                    If TypeName(values(p)) = "Integer" Or TypeName(values(p)) = "Long" Then
                        'Integer types: use {0:D} (recursive call):
                        formattedValue = Stringx.format("{0:" & specifier & "}", values(p))
                    
                    ElseIf TypeName(values(p)) = "Double" Then
                        'Non-integer types: use {0:E}
                        specifier = IIf(formatSpecifier = "G", "E", "e")
                        
                        'evaluate the exponential notation value (recursive call):
                        eNotation = Stringx.format("{0:" & specifier & "}", v)
                        
                        'get the power of eNotation:
                        ePower = Mid$(eNotation, InStr(1, UCase$(eNotation), "E-") + 1, Len(eNotation) - InStr(1, UCase$(eNotation), "E-"))
                        
                        If ePower > -5 And Abs(ePower) < precisionSpecifier Then
                            'use {0:F} when ePower > -5 and abs(ePower) < precisionSpecifier:
                            'evaluate the floating-point value (recursive call):
                             specifier = IIf(formatSpecifier = "G", "F", "f")
                             formattedValue = Stringx.format("{0:" & formatSpecifier & _
                                 IIf(precisionSpecifier <> 0, precisionString, vbNullString) & "}", values(p))
                        Else
                            'fallback to {0:E} if previous rule didn't apply:
                            formattedValue = eNotation
                        End If
                        
                    End If
                    
                    GoTo AlignFormattedValue 'Skip the "ApplyStringFormat" step, it's applied already.
                    
                    
                    Case "N", "n" 'NUMERIC format, formats string as an integer or decimal number.
                    'Precision specifier determines number of decimal digits.
                    'This implementation ignores regional settings'
                    '(hard-coded group and decimal separators).
                    
                    precisionSpecifier = CInt(precisionString)
                    If LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                        'if a non-zero precision is specified...
                        thisFormat = "#,##0"
                        thisFormat = (thisFormat & ".") & String$(precisionSpecifier, Chr$(48))
                        
                    Else 'only the "D" is specified
                        thisFormat = "#,##0"
                    End If
                    
                    
                    Case "P", "p" 'PERCENT format. Formats string as a percentage.
                    'Value is multiplied by 100 and displayed with a percent symbol.
                    'Precision specifier determines number of decimal digits.
                    
                    thisFormat = "#,##0%"
                    precisionSpecifier = CInt(precisionString)
                    If LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                        'if a non-zero precision is specified...
                        thisFormat = "#,##0"
                        thisFormat = (thisFormat & ".") & String$(precisionSpecifier, Chr$(48))
                        
                    Else 'only the "P" is specified
                        thisFormat = "#,##0"
                    End If
                    
                    'Append the percentage sign to the format string:
                    thisFormat = thisFormat & "%"
                    
                    
                    Case "R", "r" 'ROUND-TRIP format (a string that can round-trip to an identical number)
                    'example: ?StringFormat("{0:R}", 0.0000000001141596325677345362656)
                    '         ...returns "0.000000000114159632567735"
                    
                    'convert value to a Double (chop off overflow digits):
                    v = CDbl(v)
                    
                    
                    Case "X", "x" 'HEX format. Formats a string as a Hexadecimal value.
                    'Precision specifier determines number of total digits.
                    'Returned string is prefixed with "&H" to specify Hex.
                    
                    v = Hex(v)
                    precisionSpecifier = CInt(precisionString)
                    
                    If LenB(precisionString) > 0 Then 'precision here stands for left padding
                        v = Right$(String$(precisionSpecifier, "0") & v, IIf(precisionSpecifier = 0, Len(v), precisionSpecifier))
                    End If
                    
                    'add C# hex specifier, apply specified casing:
                    '(VB6 hex specifier would cause Format() to reverse the formatting):
                    v = "0x" & IIf(formatSpecifier = "X", UCase$(v), LCase$(v))
                    
                    
                    Case Else
                        Err.Raise ERR_FORMAT_EXCEPTION, _
                            ERR_SOURCE, ERR_MSG_INVALID_FORMAT_STRING
                End Select
                
                GoTo ApplyStringFormat
                
                
DateTimeFormatSpecifiers:
                Select Case formatSpecifier
                    
                    Case "c", "C" 'CUSTOM date/time format
                    'let VB Format() parse precision specifier as is:
                        thisFormat = precisionString
                    
                    Case "d" 'SHORT DATE format
                        thisFormat = "ddddd"
                        
                    Case "D" 'LONG DATE format
                        thisFormat = "dddddd"
                        
                    Case "f" 'FULL DATE format (short)
                        thisFormat = "dddddd h:mm AM/PM"
                    
                    Case "F" 'FULL DATE format (long)
                        thisFormat = "dddddd ttttt"
                    
                    Case "g"
                        thisFormat = "ddddd hh:mm AM/PM"
                        
                    Case "G"
                        thisFormat = "ddddd ttttt"
                        
                    Case "s" 'SORTABLE DATETIME format
                        thisFormat = "yyyy-mm-ddThh:mm:ss"
                    
                    Case "t" 'SHORT TIME format
                        thisFormat = "hh:mm AM/PM"
                    
                    Case "T" 'LONG TIME format
                        thisFormat = "ttttt"
                    
                    Case Else
                        Err.Raise ERR_FORMAT_EXCEPTION, _
                            ERR_SOURCE, ERR_MSG_INVALID_FORMAT_STRING
                End Select
                GoTo ApplyStringFormat
                
            End If
            
            
ApplyStringFormat:
            'apply computed format string:
            formattedValue = VBA.Strings.format(v, thisFormat)
            
            
AlignFormattedValue:
            'apply specified alignment specifier:
            If alignmentSpecifier <> vbNullString Then
            
                alignmentPadding = Abs(CInt(alignmentSpecifier))
                If CInt(alignmentSpecifier) < 0 Then
                    'negative: left-justified alignment
                    If alignmentPadding - Len(formattedValue) > 0 Then _
                        formattedValue = formattedValue & _
                            String$(alignmentPadding - Len(formattedValue), PADDING_CHAR)
                Else
                    'positive: right-justified alignment
                    If alignmentPadding - Len(formattedValue) > 0 Then _
                        formattedValue = String$(alignmentPadding - Len(formattedValue), PADDING_CHAR) & formattedValue
                End If
            End If
            
            'Replace C# hex specifier with VB6 hex specifier:
            If Stringx.contains(formattedValue, "0x", False) Then formattedValue = Replace$(formattedValue, "0x", "&H")
            
            'replace all occurrences of placeholder {i} with their formatted values:
            return_value = Replace(return_value, thisString, formattedValue, count:=1)
            
            'reset before reiterating:
            thisFormat = vbNullString
        Next
        
        
checkEscapes:
        'if there's no more backslashes, don't bother checking for the rest:
        If Not Stringx.contains(return_value, "\") Then GoTo normalExit
        
        Dim escape As EscapeSequence
        For i = 0 To escapes.count - 1
            Set escape = escapes(CStr(i))
            If Stringx.contains(return_value, escape.EscapeString, False) Then _
                return_value = Replace(return_value, escape.EscapeString, escape.ReplacementString)
        
            If Not Stringx.contains(return_value, "\") Then _
                GoTo normalExit
        Next
        
        'replace "ASCII (oct)" escape sequence
        Set regex = New RegExp
        regex.pattern = "\\(\d{3})"
        regex.ignoreCase = True
        regex.Global = True
        Set matches = regex.Execute(format_string)
        
        Dim char As Long
        If matches.count <> 0 Then
            For Each thisMatch In matches
                p = thisMatch.SubMatches(0)
                '"p" contains the octal number representing the ASCII code we're after:
                p = "&O" & p 'prepend octal prefix
                char = CLng(p)
                return_value = Replace(return_value, thisMatch.value, Chr$(char))
            Next
        End If
        
        'if there's no more backslashes, don't bother checking for the rest:
        If Not Stringx.contains(return_value, "\") Then GoTo normalExit
        
        'replace "ASCII (hex)" escape sequence
        Set regex = New RegExp
        regex.pattern = "\\x(\w{2})"
        regex.ignoreCase = True
        regex.Global = True
        Set matches = regex.Execute(format_string)
        
        If matches.count <> 0 Then
            For Each thisMatch In matches
                p = thisMatch.SubMatches(0)
                '"p" contains the hex value representing the ASCII code we're after:
                p = "&H" & p 'prepend hex prefix
                char = CLng(p)
                return_value = Replace(return_value, thisMatch.value, Chr$(char))
            Next
        End If

normalExit:
    If Stringx.contains(return_value, Chr$(27)) Then return_value = Replace(return_value, Chr$(27), "\")
    format = return_value
   
End Function

Public Function contains(ByVal text As String, ByVal searchTerm As String, Optional ByVal caseSensitive As Boolean = True) As Boolean
        
    Dim compareMethod As VbCompareMethod
    
    If caseSensitive Then
        compareMethod = vbBinaryCompare
    Else
        compareMethod = vbTextCompare
    End If
    
    contains = InStr(1, text, searchTerm, compareMethod) <> 0
    
End Function

Public Function containsAny(ByVal text As String, ByVal caseSensitive As Boolean, ParamArray searchTerms() As Variant) As Boolean
    
    Dim searchTerm As String, i As Integer, found As Boolean
    
    For i = LBound(searchTerms) To UBound(searchTerms)
        
        searchTerm = CStr(searchTerms(i))
        found = contains(text, searchTerm, caseSensitive)
        
        If found Then Exit For
    Next
    
    containsAny = found
    
End Function

Public Function startsWith(ByVal text As String, ByVal prefix As String, Optional ByVal caseSensitive As Boolean = True) As Boolean
    
    If Not caseSensitive Then
        text = LCase$(text)
        prefix = LCase$(prefix)
    End If

    startsWith = Left$(text, LenB(prefix) / 2) = prefix
End Function

Public Function endsWith(ByVal text As String, ByVal postfix As String, Optional ByVal caseSensitive As Boolean = True) As Boolean
        
    If Not caseSensitive Then
        text = LCase$(text)
        postfix = LCase$(postfix)
    End If
    
    endsWith = Right$(text, LenB(postfix) / 2) = postfix

End Function

Public Function trimChar(ByVal text As String, Optional ByVal chars As String = " " & vbCr & vbLf & vbTab & vbFormFeed & vbVerticalTab) As String
    trimChar = trimCharStart(text, chars)
    trimChar = trimCharEnd(trimChar, chars)
End Function


Public Function trimCharStart(ByVal text As String, Optional ByVal chars As String = " " & vbCr & vbLf & vbTab & vbFormFeed & vbVerticalTab) As String
    If Len(text) = 0 Then
        trimCharStart = ""
    Else
        Do While InStr(1, chars, Left(text, 1)) > 0
            text = substr(text, 1)
        Loop
        trimCharStart = text
    End If
End Function


Public Function trimCharEnd(ByVal text As String, Optional ByVal chars As String = " " & vbCr & vbLf & vbTab & vbFormFeed & vbVerticalTab) As String
    If Len(text) = 0 Then
        trimCharEnd = ""
    Else
        Do While InStr(1, chars, Right(text, 1)) > 0
            text = substr(text, 0, -1)
        Loop
        trimCharEnd = text
    End If
End Function


Public Function substr(ByVal text As String, ByVal startIndex As Integer, Optional ByVal length As Variant) As String
    If startIndex < 0 Then
        startIndex = Len(text) + startIndex
    End If
    
    Dim intLength As Integer
    If IsMissing(length) Then
        intLength = Len(text) - startIndex
    Else
        intLength = length
    End If
    
    If intLength < 0 Then
        intLength = Len(text) - startIndex + intLength
    End If
    
    If intLength < 0 Then
        Err.Raise E_INDEXOUTOFRANGE, "Stringx.substr()", "End index is smaller than start index."
    End If
    
    substr = Mid(text, startIndex + 1, intLength)
End Function


Public Function split(text As String, Optional separator As String = " ") As List
    If Len(text) = 0 Then
        Set split = List_createT("String")
        split.push ""
    ElseIf separator = "" Then
        Set split = toChars(text)
    Else
        Dim arr() As String
        arr = VBA.Strings.split(text, separator)
        
        Set split = List_createT("String")
        split.append arr
    End If
End Function


Public Function join(l As List, separator As String) As String
    Dim result As String
    result = ""
    
    Dim entry As Variant
    For Each entry In l
        If IsNumeric(entry) Or varType(entry) = vbString Then
            result = result & entry
        Else
            ' If we don't have a type that nicely stringifies, just use the gist.
            result = result & Variants.gist(entry)
        End If
        
        result = result & separator
    Next
    
    If result <> "" Then
        join = Stringx.substr(result, 0, -Len(separator))
    Else
        join = result
    End If
End Function


Public Function toChars(text As String) As List
    Dim l As List
    Set l = List_createT("String")
    
    Dim runner As Integer
    For runner = 0 To Len(text) - 1
        l.push substr(text, runner, 1)
    Next
    Set toChars = l
End Function


Public Function repeat(text As String, times As Integer) As String
    If times < 0 Then
        Err.Raise E_ARGUMENTOUTOFRANGE, "Stringx.Repeat()", "Repetition count must not be negative."
    End If

    Dim result As String
    result = ""
    
    Dim runner As Integer
    For runner = 1 To times
        result = result & text
    Next
    
    repeat = result
End Function

'This function is a expansion of replace
'Parameter:
'           input:      String,     contrains the string to edit
'           values():   ParamArray, contains the string/chars which have to be replaced
'Return:
'           trans:      String,     contains the changed input string
Public Function trans(inputString As String, ParamArray values()) As String
 
    Dim CountParams     As Integer
    Dim Counter         As Integer
    
    Dim RowValue        As Variant
    Dim CheckError      As Variant
    
    Dim toChange        As List
    Dim toApply         As List

    Dim Replaced        As Boolean
    
    CountParams = UBound(values) - LBound(values) + 1
    
    If CountParams Mod 2 <> 0 Then
        Err.Raise E_INVALIDINPUT
    End If
        
    Counter = 0

    For Each RowValue In values()
                
        If Counter Mod 2 = 0 Then

            If varType(RowValue) <> vbObject Or TypeName(RowValue) <> "List" Then
                Set toChange = List_create(RowValue)
            Else
                Set toChange = RowValue
            End If
            
            For Each CheckError In toChange
                If CheckError = "" Then
                    Err.Raise E_INVALIDINPUT
                End If
            Next
            
        ElseIf Counter Mod 2 <> 0 Then
        
            If varType(RowValue) <> vbObject Or TypeName(RowValue) <> "List" Then
                Set toApply = List_create(RowValue)
                
               ReplaceStringWithLists inputString, toChange, toApply
                                 
            Else
            
                Set toApply = RowValue
                                
                If toApply.elems <> toChange.elems Then
                    Err.Raise E_INVALIDINPUT
                End If
                                
                ReplaceStringWithLists inputString, toChange, toApply
                
            End If
        End If
        
        Counter = Counter + 1
    Next
    trans = inputString
End Function


Private Function ReplaceStringWithLists(ByRef inputString As String, ByRef toChange As List, ByRef toApply As List) As Boolean
    On Error GoTo Error
    ReplaceStringWithLists = False
    
    Dim sReplace        As Variant
    Dim Fealding        As Variant
    Dim ReplacedBefore  As String
    
    If toApply.elems < toChange.elems Then
        
        For Each sReplace In toChange
            For Each Fealding In toApply
                If ReplacedBefore <> sReplace Then
                    inputString = Replace(inputString, sReplace, Fealding)
                    
                    ReplacedBefore = sReplace
                    Exit For
                Else
                    GoTo NextReplacement_sReplace
                End If
NextReplacement_sReplace:
            Next
        Next
        
    Else
    
        For Each Fealding In toApply
            For Each sReplace In toChange
                If ReplacedBefore <> sReplace Then
                    inputString = Replace(inputString, sReplace, Fealding)
                    
                    ReplacedBefore = sReplace
                    Exit For
                Else
                    GoTo NextReplacement_Fealding
                End If
NextReplacement_Fealding:
            Next
        Next
    
    End If
    
    ReplacedBefore = ""
    
    ReplaceStringWithLists = True

Error:
    ReplaceStringWithLists = False
End Function

Private Sub Class_Initialize()
    Dim factory As New EscapeSequence
    escapes.Add factory.Create("\n", vbNewLine), "0"
    escapes.Add factory.Create("\q", Chr$(34)), "1"
    escapes.Add factory.Create("\t", vbTab), "2"
    escapes.Add factory.Create("\a", Chr$(7)), "3"
    escapes.Add factory.Create("\b", Chr$(8)), "4"
    escapes.Add factory.Create("\v", Chr$(13)), "5"
    escapes.Add factory.Create("\f", Chr$(14)), "6"
    escapes.Add factory.Create("\r", Chr$(15)), "7"
    Set factory = Nothing
End Sub
