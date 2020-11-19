VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XlUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Function nameExists(rangeName As String, Optional wb As Workbook) As Boolean
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim foo As name
    On Error Resume Next
    Set foo = wb.Names(rangeName)
    nameExists = Err.number = 0
    On Error GoTo 0
End Function

Public Function sheetExists(sheetName As String, Optional wb As Workbook) As Boolean
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = wb.Sheets(sheetName)
    On Error GoTo 0
    
    sheetExists = Not sheet Is Nothing
End Function


Public Function commandBarExists(cmdBarName As String) As Boolean
    Dim cmdBar As CommandBar
    On Error Resume Next
    Set cmdBar = Application.CommandBars(cmdBarName)
    On Error GoTo 0
    
    commandBarExists = Not cmdBar Is Nothing
End Function


Public Function openWorkbookMaybeOpen(path As String) As Workbook
' Works only with absolute paths.
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.FullName = path Then
            Set openWorkbookMaybeOpen = wb
            Exit Function
        End If
    Next
    
    Set openWorkbookMaybeOpen = Application.Workbooks.Open(filename:=path)
End Function


Public Function isWorkbookOpen(path As String) As Boolean
' Works only with absolute paths.
    isWorkbookOpen = False
    
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.FullName = path Then
            isWorkbookOpen = True
            Exit Function
        End If
    Next
End Function

Public Function isWorkbookNameOpen(workbookFilename As String) As Boolean
    isWorkbookNameOpen = False
    
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.name = workbookFilename Then
            isWorkbookNameOpen = True
            Exit Function
        End If
    Next
End Function

Public Function getOpenWorkbook(path As String) As Workbook
' Works only with absolute paths.
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.FullName = path Then
            Set getOpenWorkbook = wb
            Exit Function
        End If
    Next
    
    Err.Raise E_WORKBOOKNOTOPEN, "XlUtils.getOpenWorkbook", "The given workbook is not open."
End Function

'Public Sub makeWritable(sheet As Worksheet)
'    If sheet.ProtectContents Then
'        sheet.Unprotect
'        sheet.Protect UserInterfaceOnly:=True
'    End If
'End Sub

'Public Sub makeWritable2(sheet As String, Optional wb As Workbook)
'    Dim myWb As Workbook
'    If wb Is Nothing Then
'        Set myWb = Application.ThisWorkbook
'    Else
'        Set myWb = wb
'    End If
'    makeWritable2 myWb.Sheets(sheet)
'End Sub

Public Function rowColToExcel(ByVal row As Integer, ByVal column As Integer) As String
    rowColToExcel = numToLetters(column) & row
End Function

Public Function numToLetters(ByVal num As Integer) As String
    Static alphabet As String
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    If num <= 0 Then
        Err.Raise E_INDEXOUTOFRANGE, "numToLetters()", "Column must be greater than 0."
    End If
    
    Dim result As String
    result = ""
    
    Dim remainder As Integer
    
    ' We have to work with 0-based numbers (0=A, 25=Z), otherwise the modulo operator won't work.
    ' (26 / 26 = 1 Remainder 0 and not 0 Remainder 26 which would lead to A0 and not Z)
    num = num - 1
    
    While num >= 0
        remainder = num Mod Len(alphabet)
        ' + 1 because Mid takes 1-based indexes
        result = Mid(alphabet, remainder + 1, 1) & result
        
        ' - 1, to get back to a 0-based system.
        ' 1 Remainder 0 should result in A(=0).
        ' 0 Remainder 0 should trigger loop end.
        num = num \ Len(alphabet) - 1
    Wend
    
    numToLetters = result
End Function

Public Function lettersToNum(letters As String) As Integer
    Static alphabet As String
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    Dim result As Integer: result = 0

    Dim parts As List: Set parts = Stringx.toChars(letters)
    Dim part As Variant
    For Each part In parts
        Dim value As Integer: value = InStr(1, alphabet, part, vbTextCompare)
        result = result * Len(alphabet) + value
    Next
    lettersToNum = result
End Function

