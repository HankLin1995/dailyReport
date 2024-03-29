Attribute VB_Name = "GIT"
'TODO:Export folder need to be killed

Sub ExportCodesToFolder()

'Type: 1=bas,2=cls,3=frm

myFolder = getSavedFolder

Call killFilesInFolder(myFolder)

Set VBProj = ThisWorkbook.VBProject
For Each VBComp In VBProj.VBComponents
    
    Select Case VBComp.Type
    
        Case 1: myExtension = ".bas"
        Case 2: myExtension = ".cls"
        Case 3: myExtension = ".frm"
        
        Case 100: myExtension = ".doccls"
    
    End Select
    
    full_path = myFolder & "\" & VBComp.Name & myExtension
    
    If myExtension <> "" Then
    
        VBComp.Export (full_path)
    
    End If
    
    If myExtension = ".doccls" And CountFileLines(full_path) = 9 Then Kill full_path
    
Next VBComp
    
End Sub

Sub killFilesInFolder(folderPath)

Set coll_path = GetFilePathsInFolder(folderPath)

For Each FilePath In coll_path

    FileName = mid(FilePath, InStrRev(FilePath, "\") + 1)
    fileExtension = mid(FileName, InStrRev(FileName, ".") + 1)
    
    If fileExtension = "frm" Or fileExtension = "bas" Or fileExtension = "cls" Or fileExtension = "doccls" Then
        Kill FilePath
    End If
Next

End Sub

Sub ImportCodes()

myFolder = getSavedFolder

Set coll_path = GetFilePathsInFolder(myFolder)

For Each FilePath In coll_path

    FileName = mid(FilePath, InStrRev(FilePath, "\") + 1)
    fileExtension = mid(FileName, InStrRev(FileName, ".") + 1)
    
    If fileExtension = "frm" Or fileExtension = "bas" Or fileExtension = "cls" Then
        Call ImportCode(FilePath, FileName)
    End If

Next

End Sub

Sub ImportCode(ByVal FilePath As String, ByVal FileName As String)

extension = mid(FileName, InStrRev(FileName, ".") + 1)
CodeName = mid(FileName, 1, InStrRev(FileName, ".") - 1)

If CodeName = "GIT" Then Exit Sub

Set VBProj = ThisWorkbook.VBProject

'If checkIfCodeExist(CodeName) = True Then
'
'    Set vbcomp = VBProj.VBComponents(CodeName)
'    VBProj.VBComponents.Remove (vbcomp)
'
'End If

VBProj.VBComponents.Import (FilePath)

End Sub

Sub DeleteCodes()

'Type: 1=bas,2=cls,3=frm

Set VBProj = ThisWorkbook.VBProject
For Each VBComp In VBProj.VBComponents
    
    Select Case VBComp.Type
    
        Case 1: myExtension = ".bas"
        Case 2: myExtension = ".cls"
        Case 3: myExtension = ".frm"
        
        Case 100: myExtension = ".doccls"
    
    End Select
    
    If VBComp.Type <> 100 And VBComp.Name <> "GIT" Then

        VBProj.VBComponents.Remove (VBComp)
        
    End If
    
Next VBComp

End Sub

'--------FUNCTION------------

Function GetFilePathsInFolder(ByVal folderPath As String)

    Dim coll As New Collection

    Dim fso As Object
    'Dim folderPath As String
    Dim folder As Object
    Dim file As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

   ' folderPath = getSavedFolder
    Set folder = fso.GetFolder(folderPath)
    
    For Each file In folder.Files

        coll.Add file.Path
        
    Next file
    
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
    
    Set GetFilePathsInFolder = coll
    
End Function

Function getSavedFolder()

    Set fldr = Application.FileDialog(4)
    
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        '.InitialFileName = getThisWorkbookPath
        If .Show = -1 Then FolderName = .SelectedItems(1)
    End With
getSavedFolder = FolderName

End Function

Function checkIfCodeExist(ByVal checkName As String) 'useless

Set VBProj = ThisWorkbook.VBProject
Set VBComps = VBProj.VBComponents

checkIfCodeExist = False

For Each it In VBComps

    If it.Name = checkName Then
        
        checkIfCodeExist = True: Exit Function
        
    End If
Next

End Function

Function CountFileLines(ByVal FilePath)

    Dim fileContent As String
    Dim FileNumber As Integer
    Dim lineCount As Long
    
    ' Open the text file
    FileNumber = FreeFile
    Open FilePath For Input As FileNumber
    
    ' Read the file content line by line and count the lines
    Do Until EOF(FileNumber)
        Line Input #FileNumber, fileContent
        lineCount = lineCount + 1
    Loop
    
    ' Close the file
    Close FileNumber
    
    ' Display the line count in cell A1
    CountFileLines = lineCount
    
End Function

'--------TMP_CODE-------------

Function tmp_deleteCodes()

Set VBProj = ThisWorkbook.VBProject
Set VBComps = VBProj.VBComponents

For Each it In VBComps
    
    If it.Name Like "*2" And it.Type <> 100 Then

        CodeName = it.Name
        
        Set VBComp = VBProj.VBComponents(CodeName)
        VBProj.VBComponents.Remove (VBComp)
        
    End If
    
Next

End Function



