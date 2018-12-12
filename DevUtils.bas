Attribute VB_Name = "DevUtils"
Public Sub importModules(repoPath As String, Optional wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    addReferenceInternalXZY "{0D452EE1-E08F-101A-852E-02608C4D0BB4}", "Microsoft Forms 2.0 Object Library", wb
    addReferenceInternalXZY "{3F4DACA7-160D-11D2-A8E9-00104B365C9F}", "Microsoft VBScript Regular Expressions 5.5", wb, 5, 5
    addReferenceInternalXZY "{0002E157-0000-0000-C000-000000000046}", "Microsoft Visual Basic for Applications Extensibility 5.3", wb
    addReferenceInternalXZY "{420B2830-E718-11CF-893D-00A0C9054228}", "Microsoft Scripting Runtime", wb
    addReferenceInternalXZY "{662901FC-6951-4854-9EB2-D9A2570F2B2E}", "Microsoft WinHTTP Services, version 5.1", wb
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim repoFolder As Object: Set repoFolder = fso.GetFolder(repoPath)

    Dim vbaFile As Object
    For Each vbaFile In repoFolder.Files
        importModuleStuffs vbaFile.path, wb
    Next
    
    For Each vbaFile In repoFolder.SubFolders("tests").Files
        importModuleStuffs vbaFile.path, wb
    Next
End Sub

Private Sub importModuleStuffs(path As String, wb As Workbook)
    If Right$(path, 12) = "DevUtils.bas" Then
        Exit Sub
    End If
    
    If LCase$(Right$(path, 4)) = ".bas" _
    Or LCase$(Right$(path, 4)) = ".cls" _
    Or LCase$(Right$(path, 4)) = ".frm" Then
        On Error Resume Next
        wb.VBProject.VBComponents.Import path
        If Err.number <> 0 Then
            Debug.Print "Failed to import " & path
        End If
        On Error GoTo 0
    End If
End Sub

Public Sub exportModules(repoPath As String, Optional wb As Workbook)
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    If Right$(repoPath, 1) <> "/" Then
        repoPath = repoPath & "/"
    End If
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(repoPath) Then
        fso.CreateFolder repoPath
    End If
    If Not fso.FolderExists(repoPath & "tests/") Then
        fso.CreateFolder repoPath & "tests/"
    End If
    
    Dim component As Object
    For Each component In wb.VBProject.VBComponents
        Dim ext As String
        Select Case component.Type
        Case vbext_ct_MSForm
            ext = ".frm"
        Case vbext_ct_StdModule
            ext = ".bas"
        Case vbext_ct_ClassModule
            ext = ".cls"
        End Select
        
        If ext <> "" Then
            If LCase$(Left$(component.name, 5)) = "test_" Then
                component.Export filename:=repoPath & "tests/" & component.name & ext
            Else
                component.Export filename:=repoPath & component.name & ext
            End If
        End If
    Next
End Sub

Private Sub addReferenceInternalXZY(guid As String, name As String, wb As Workbook, Optional majorVersion As Integer = 1, Optional minorVersion As Integer = 0)
    On Error Resume Next
    Err.clear
    
    wb.VBProject.References.AddFromGuid guid:=guid, major:=majorVersion, minor:=minorVersion
    
    Select Case Err.number
    Case Is = 32813
        ' Is already there -> Ignore.
    Case Is = vbNullString
        ' All good.
    Case Else
        MsgBox "Adding the reference " & name & " failed. Do it manually.", vbCritical + vbOKOnly, "Verweis hinzufügen"
    End Select
    
    On Error GoTo 0
End Sub
