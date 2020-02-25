Attribute VB_Name = "Exception"
Option Explicit

' The error numbers encode an error group and an error number.
' Every group has 128 possible error numbers (0-127).
' 0000 0000 0000 0000 0000 0000 0111 1111 = &H 00 00 00 7F =        127 = error number mask
' 0111 1111 1111 1111 1111 1111 1000 0000 = &H 7F FF FF 80 = 2147483520 = group mask
' Error numbers <= 512 are reserved by Microsoft.
' Not yet sure whether this is overkill...

Public Const errNumberMask = &H7F
Public Const errGroupMask = &H7FFFFF80

' Microsoft reserved error codes
' 0 - 512
' https://msdn.microsoft.com/en-us/library/ms234761(v=vs.90).aspx
Public Const E_INDEXOUTOFRANGE = 9

' Custom codes
' Common errors. 513 - 639 (512 is the mask, but 512 is still MS reserved)
Public Const E_COMMONERRORS = 512
Public Const E_ARGUMENTOUTOFRANGE = 513
Public Const E_ILLEGALSTATE = 514
Public Const E_INTERNALERROR = 515
Public Const E_TYPEMISMATCH = 516
Public Const E_INVALIDINPUT = 517

' CalculateRequestParser errors. 640 - 767
Public Const E_CALCULATEREQUESTPARSER = 640
Public Const E_DUPLICATEINPUT = 640

' XlUtils errors. 768 - 895
Public Const E_XLUTILS = 768
Public Const E_WORKBOOKNOTOPEN = 768

' IO errors. 896 - 1023
Public Const E_IO = 896
Public Const E_FILEEXISTS = 896
Public Const E_FILENOTFOUND = 897
Public Const E_UNKNOWNENCODING = 898


Public Function isErrGroup(errNo As Long, errGroup As Long) As Boolean
    isErrGroup = getErrGroup(errNo) = errGroup
End Function

Public Function getErrGroup(errNo As Long) As Long
    getErrGroup = errNo And errGroupMask
End Function

Public Function errNoToString(errNo As Integer) As String
    Select Case errNo
    ' Microsoft reserved
    Case 9: errNoToString = "E_INDEXOUTOFRANGE"
    
    'Common errors
    Case 513: errNoToString = "E_ARGUMENTOUTOFRANGE"
    Case 514: errNoToString = "E_ILLEGALSTATE"
    Case 515: errNoToString = "E_INTERNALERROR"
    Case 516: errNoToString = "E_TYPEMISMATCH"
    Case 517: errNoToString = "E_INVALIDINPUT"
    
    ' CalculateRequestParser errors
    Case 640: errNoToString = "E_DUPLICATEINPUT"
    
    ' XlUtils errors
    Case 768: errNoToString = "E_WORKBOOKNOTOPEN"
    
    ' IO errors
    Case 896: errNoToString = "E_FILEEXISTS"
    Case 897: errNoToString = "E_FILENOTFOUND"
    Case 898: errNoToString = "E_UNKNOWNENCODING"
    
    Case Else: errNoToString = ""
    End Select
End Function

Public Sub rethrow()
    Err.Raise Err.number, Err.source, Err.description, Err.HelpFile, Err.HelpContext
End Sub

Public Function exceptionToString(Optional er As ErrObject) As String
    If er is Nothing Then
        Set er = Err
    End If
    
    exceptionToString = "EX "
    exceptionToString = exceptionToString & Err.number
    
    If errNoToString(Err.number) <> "" Then
        exceptionToString = exceptionToString & " " & errNoToString(Err.number)
    End If
    
    exceptionToString = exceptionToString & ": "
    
    If Err.description <> "" Then
        exceptionToString = exceptionToString & Err.description & vbCrLf
    Else
        exceptionToString = exceptionToString & "And exception occurred." & vbCrLf
    End If
    
    If Err.source <> "" Then
        exceptionToString = exceptionToString & "Source: " & Err.source & vbCrLf
    End If
    
    If Err.HelpFile <> "" Then
        exceptionToString = exceptionToString & "HelpFile: " & Err.HelpFile & vbCrLf
    End If
    
    If Err.HelpContext > 0 Then
        exceptionToString = exceptionToString & "HelpContext: " & Err.HelpContext & vbCrLf
    End If
End Function

Public Sub printException(Optional er As ErrObject)
    If er is Nothing Then
        Set er = Err
    End If
    
    Debug.Print exceptionToString(er)
End Sub

Public Sub msgBoxException(Optional er As ErrObject)
    If er Is Nothing Then
        Set er = Err
    End If
    
    If er.HelpFile <> "" Then
        MsgBox exceptionToString(er), vbMsgBoxHelpButton, Title:="Exception", HelpFile:=er.HelpFile, Context:=er.HelpContext
    Else
        MsgBox exceptionToString(er), Title:="Exception"
    End If
End Sub
