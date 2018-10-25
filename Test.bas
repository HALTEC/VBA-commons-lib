Attribute VB_Name = "Test"
Option Explicit

Private test_counter As Integer
Private testGroups As List
Private silent As Boolean

Private Sub addTest(pass As Boolean, testType As String, errorInfo As String, message As Variant)
    If testGroups.elems = 0 Then
        Err.Raise E_ILLEGALSTATE, "Test.addTest()", "Need to call gStart() before doing a test."
    End If
    
    testGroups.last.addTest pass, testType, errorInfo, CStr(message)
    
    test_counter = test_counter + 1
End Sub

Public Sub ok(pass As Boolean, Optional message As Variant)
    addTest pass, "ok", "", IIf(IsMissing(message), "", CStr(message))
End Sub


Public Sub equals(ByVal value As Variant, ByVal expected As Variant, Optional message As Variant)
    Dim areEqual As Boolean
    
    If varType(value) = vbObject <> varType(expected) = vbObject Then
        areEqual = False
        GoTo Output
    End If
    
    If varType(value) = vbObject And TypeName(value) <> TypeName(expected) Then
        areEqual = False
        GoTo Output
    End If
    
    If IsNumeric(value) <> IsNumeric(expected) Then
        areEqual = False
        GoTo Output
    End If
    
    If varType(value) = vbEmpty Or varType(value) = vbNull Then
        areEqual = True
        GoTo Output
    End If
    
    ' Array handling
    If CBool(varType(value) And vbArray) Then
        If LBound(value) <> LBound(expected) Or UBound(value) <> UBound(expected) Then
            areEqual = False
            GoTo Output
        End If
        
        Dim runner As Integer
        For runner = LBound(value) To UBound(value)
            If value(runner) <> expected(runner) Then
                areEqual = False
                GoTo Output
            End If
        Next
        
        areEqual = True
        GoTo Output
    End If
    
    ' Nothing handling
    If varType(value) = vbObject Then
        If value Is Nothing Or expected Is Nothing Then
            areEqual = (value Is Nothing) = (expected Is Nothing)
            GoTo Output
        End If
    End If
    
    ' Special case for List and Map.
    If varType(value) = vbObject Then
        If (TypeName(value) = "List" And TypeName(expected) = "List") _
                Or (TypeName(value) = "Map" And TypeName(expected) = "Map") Then
            Dim tmp As Object
            Set tmp = expected
            areEqual = value.equals(tmp)
            GoTo Output
        End If
    End If

    ' Default to comparison via =
    areEqual = value = expected
    GoTo Output

Output:
    addTest areEqual, "equals", "found: " & Variants.gist(value) & ", expected: " & Variants.gist(expected), IIf(IsMissing(message), "", CStr(message))
    
End Sub


Public Sub checkError(ByVal errNo As Long, Optional message As Variant)
    addTest Err.number = errNo, "checkError", "found: " & Err.number & ", expected: " & errNo, IIf(IsMissing(message), "", CStr(message))
End Sub

Public Sub checkNoError(Optional message As Variant)
    addTest Err.number = 0, "checkNoError", "ErrorNo: " & Err.number, IIf(IsMissing(message), "", CStr(message))
End Sub

Public Sub fail(Optional message As Variant)
    addTest False, "fail", "", IIf(IsMissing(message), "", CStr(message))
End Sub


Public Sub pass(Optional message As Variant)
    addTest True, "pass", "", IIf(IsMissing(message), "", CStr(message))
End Sub


Public Sub gStart(name As String)
    If testGroups Is Nothing Then
        Set testGroups = List_createT("TestGroup")
    End If
    
    testGroups.push TestGroup_create(name)
End Sub


Public Sub gStop()
    If testGroups Is Nothing Then
        Err.Raise E_ILLEGALSTATE, "Test.gStop()", "Can't call gStop() without calling gStart() first."
    End If
    
    If testGroups.elems = 0 Then
        Err.Raise E_INTERNALERROR, "Test.gStop()", "Found a testGroups list but no entries."
    End If
    
    ' Pop
    Dim entry As TestGroup: Set entry = testGroups.pop
    
    If testGroups.elems > 0 Then
        testGroups.last.addGroup entry
    Else
        Debug.Print entry.getError
        Debug.Print "Test count: " & vbTab & test_counter
        Debug.Print "Failure count: " & vbTab & entry.errorCount
        If entry.errorCount > 0 Then
            Debug.Print "!!! Failure !!!!"
        Else
            Debug.Print "Success"
        End If
        test_counter = 0
    End If
    
End Sub

