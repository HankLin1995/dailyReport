Attribute VB_Name = "test"

Sub test_collectFilePaths()

Dim f As New clsMyfunction
Dim coll As New Collection

With Sheets("Check")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 3 To lr
    
        MyCode = .Cells(r, 2)
        myNum = .Cells(r, 3)
    
        myPath = ThisWorkbook.Path & "\��d��PDF\" & MyCode & "-" & myNum & ".pdf"
    
        If f.IsFileExists(myPath) = True Then
            
            coll.Add myPath
        
        End If
        
        myPath = ThisWorkbook.Path & "\�d��Ӥ�Output\" & MyCode & "-" & myNum & ".pdf"
    
        If f.IsFileExists(myPath) = True Then
        
            coll.Add myPath
        
        End If
    Next

End With

Call WriteCollectionToTxt(coll)

myTime = Now()

Call RunTestGPTWithParameters

Call CheckFileCreateTimeWithRetryAndDelay(myTime)

End Sub

Sub RunTestGPTWithParameters()

    Dim path1 As String
    Dim myret As Long
    Dim command As String
    Dim searchKeyword As String
    Dim savePath As String

    ' �]�m�i������Ҧb���ؿ�
    'path1 = ThisWorkbook.Path ' ThisWorkbook.Path & "\GooglePhotoDownloader"
    'ChDir path1

    PDF_PATH = ThisWorkbook.Path & "\Lib\Merge\merge.pdf"
    TXT_PATH = ThisWorkbook.Path & "\Lib\Merge\file_with_paths.txt"
    
    command = ThisWorkbook.Path & "\Lib\Merge\Merge.exe """ & TXT_PATH & """ """ & PDF_PATH & """"
    
    ' �ϥ�Shell��ƹB��i������õ��ݨ䧹��
    myret = Shell(command, 1)

End Sub

Sub CheckFileCreateTimeWithRetryAndDelay(ByVal CurrentTime)
    ' ����?
    FilePath = ThisWorkbook.Path & "\Lib\merge\merge.pdf"
    
    ' ?�m�e�Ԫ��t��??�]?��?�m?1�p?�^
    ToleranceHours = 0 '1 / 24
    
    ' �̤j??��?
    MaxTries = 20
    
    Dim FileCreateTime As Date
    'Dim CurrentTime As Date
    Dim Tries As Integer
    
    Tries = 1
    
    Do While Tries <= MaxTries
    
        Debug.Print Tries
    
        If Dir(FilePath) = "" Then
            ' ��󤣦s�b�A����2��?�A??
            Tries = Tries + 1
            If Tries <= MaxTries Then
                Application.Wait Now + TimeValue("00:00:02")
            End If
        Else
            ' ���s�b�A?�����?��??
            FileCreateTime = FileDateTime(FilePath)
            
            ' ?��?�e??
            'CurrentTime = Now
            
            ' ��?���?��??�M?�e??�A��?�e��??
            If FileCreateTime > CurrentTime Then
                Shell "cmd /c start " & FilePath, vbNormalFocus
                Exit Do
            Else
                ' ���?��??�b�e�ԭS?�~�A����2��?�M�Z��?
                Tries = Tries + 1
                If Tries <= MaxTries Then
                    Application.Wait Now + TimeValue("00:00:02")
                End If
            End If
        End If
    Loop
    
    If Tries > MaxTries Then
        MsgBox "�w?��̤j��?��?�C"
    End If
End Sub




Sub WriteCollectionToTxt(ByVal coll)
    'Dim col As Collection
    Dim Item As Variant
    Dim FilePath As String
    Dim FileName As String
    Dim FileNumber As Integer
    
    ' ?�ؤ@?�s�����X�}�K�[?�u�]?���u�O�ܨҡA�A�ݭn�ۤv��R�A�����X�^
    'Set col = New Collection
    'col.Add "Item1"
    'col.Add "Item2"
    'col.Add "Item3"
    
    ' ���w���W�M��?
    FileName = "file_with_paths.txt"
    FilePath = ThisWorkbook.Path & "\Lib\Merge\" & FileName
    
    ' ��?�奻���H?��?�J
    FileNumber = FreeFile
    Open FilePath For Output As FileNumber
    
    ' �M?���X�}?�C??��?�J�奻���
    For Each Item In coll
        Print #FileNumber, Item
    Next Item
    
    ' ??���
    Close FileNumber
End Sub


Sub test_CopyFileToFolder()
    Dim SourceFilePath As String
    Dim DestinationFolder As String
    Dim SourceFileName As String
    Dim NewFilePath As String
    
    With Sheets("Check")
    
        For Each rng In Selection
    
            r = rng.Row
    
            If r > 2 Then
            
                If .Cells(r, 2) <> "" Then
                
                    newFileName = .Cells(r, 2) & "-" & .Cells(r, 3)
                
                End If
            
            End If
    
            Exit For
            
        Next
    
    End With
    
    If newFileName = "" Then MsgBox "�Х��ؿ�n�k�ɪ���m!", vbCritical

    ' ���ϥΪ̿�ܤ@���ɮ�
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "��ܭn�ƻs���ɮ�"
        If .Show = -1 Then
            SourceFilePath = .SelectedItems(1)
        Else
            'MsgBox "������ɮסC�ާ@�w�����C"
            Exit Sub
        End If
    End With

    ' ���w�ؼи�Ƨ�
    DestinationFolder = ThisWorkbook.Path & "\��d��PDF\"

    ' ��o��l�ɮצW��
    SourceFileName = mid(SourceFilePath, InStrRev(SourceFilePath, "\") + 1)

    ' �զX�s���ɮ׸��|
    NewFilePath = DestinationFolder & newFileName & ".pdf" 'SourceFileName

    ' �ˬd�ؼи�Ƨ��O�_�s�b�A�p���s�b�h�Ы�
    If Dir(DestinationFolder, vbDirectory) = "" Then
        MkDir DestinationFolder
    End If

    ' �ƻs�ɮ�
    FileCopy SourceFilePath, NewFilePath

    ' �ˬd�ɮ׬O�_���\�ƻs
    
    With Sheets("Check")
    
        If Dir(NewFilePath) <> "" Then
            'MsgBox "�ɮפw���\�ƻs�� " & NewFilePath
            .Cells(r, "H") = "V" 'NewFilePath
            
            On Error Resume Next
            .Cells(r, "H").Comment.Delete
            On Error GoTo 0
            
            .Cells(r, "H").AddComment
            .Cells(r, "H").Comment.Text Text:=NewFilePath
    
        End If
        
    End With
    
End Sub


Sub test_getTestItems()

Dim coll_tests As New Collection

With Sheets("Budget")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

    For r = 3 To lr
    
        pcces_item = .Cells(r, 2)
        
        If pcces_item Like "*����W�d�μз�*" Then
            
            If pcces_item Like "*;*" Then MsgBox pcces_item & vbNewLine & "�t���D�k�r�ši;�j�A�Эץ�!", vbCritical: End
            
            pcces_unit = .Cells(r, 3)
            pcces_amount = .Cells(r, 4)
            
            test_name = getTestReplaceName(pcces_item)
        
            If test_name <> "" Then
                
                coll_tests.Add test_name & ";" & pcces_amount & ";" & pcces_unit
                
            End If
            
        End If
    
    Next

End With

Call extractToMainTest(coll_tests)

Call cmdResetReport

End Sub

Sub extractToMainTest(ByVal coll)

Call ReturnMainRow(arr)

With Sheets("Main")

    .Cells(arr(1) + 1, 1).Resize(arr(2) - arr(1) - 1, 5).ClearContents

    For i = 1 To coll.Count
    
        tmp = split(coll(i), ";")
    
        r = arr(1) + i
    
        .Cells(r, 1) = tmp(0)
        .Cells(r, 2) = tmp(1)
        .Cells(r, 3) = tmp(2)
    
    Next

End With

End Sub

Function getTestReplaceName(ByVal test_item As String)

Dim f As New clsMyfunction

shtName = "TestReplace"

Set coll_rows = f.getRowsByUser(shtName, "A", test_item)

If coll_rows.Count = 0 Then

        getTestReplaceName = InputBox(test_item & ":���w�q�O�W!" & vbNewLine & "�п�J�O�W:", , test_item)
        Call f.AppendData(shtName, Array(test_item, getTestReplaceName))
        Exit Function

End If

For Each r In coll_rows

    Set rng = Sheets(shtName).Cells(r, 1)
    
    r = rng.Row
    getTestReplaceName = rng.Offset(0, 1).Value

    Exit For 'only deal one time
    
Next


End Function

Sub importOldData()

'getfile

Set wb = Workbooks(getWorkbookName)

Debug.Print wb.Name

'collectMainInf

With wb.Sheets("Main")

Set rng_last = .Cells.SpecialCells(xlCellTypeLastCell)

Set rngs = .Range("B1:C6")
rngs.Copy ThisWorkbook.Sheets("Main").Range("B1:C6")

'collectTestInf

Set rngs = .Range("A14:D" & rng_last.Row)
rngs.Copy ThisWorkbook.Sheets("Main").Range("A10")

'collectReportInf
Set rngs = .Range("F1:" & rng_last.Address)
rngs.Copy ThisWorkbook.Sheets("Main").Range("F1")

End With

With wb.Sheets("Budget")

'collectRecords
Set rng_last = .Cells.SpecialCells(xlCellTypeLastCell)
Set rngs = .Range("A2:" & rng_last.Address)
rngs.Copy ThisWorkbook.Sheets("Budget").Range("A2")

End With

With wb.Sheets("Records")

'collectRecords
Set rng_last = .Cells.SpecialCells(xlCellTypeLastCell)
Set rngs = .Range("A2:" & rng_last.Address)
rngs.Copy ThisWorkbook.Sheets("Records").Range("A2")

End With
'collectDiary
With wb.Sheets("Diary")

Set rng_last = .Cells.SpecialCells(xlCellTypeLastCell)
Set rngs = .Range("A2:" & rng_last.Address)
rngs.Copy ThisWorkbook.Sheets("Diary").Range("A2")

End With
'collectMix
With wb.Sheets("Mix")

Set rng_last = .Cells.SpecialCells(xlCellTypeLastCell)
Set rngs = .Range("A2:" & rng_last.Address)
rngs.Copy ThisWorkbook.Sheets("Mix").Range("A2")

End With

End Sub

Function getWorkbookName()

Dim coll As New Collection

For Each wb In Workbooks

    If wb.Name <> ThisWorkbook.Name Then
    
    coll.Add wb.Name
    
    End If

Next

If coll.Count = 0 Then MsgBox "�Х��}�ҭn�פJ������!", vbCritical: End

For i = 1 To coll.Count

    p = p & i & "." & coll(i) & vbNewLine

Next

If coll.Count = 1 Then

getWorkbookName = coll(1)

Else

myIndex = InputBox("�п�J�n�פJ���ɮצW��" & vbNewLine & p)

getWorkbookName = coll(myIndex)

End If

End Function































Function splitPhotoList(ByVal s As String)

If s = "" Then Exit Function

s = "G:\�ڪ����ݵw��\ExcelVBA\�ʳy�����\@�ʳy�����DEV\�W�ү���\�I�u�Ӥ�\1120130-���O�Υ��]��d\IMG_3824.JPG>1,G:\�ڪ����ݵw��\ExcelVBA\�ʳy�����\@�ʳy�����DEV\�W�ү���\�I�u�Ӥ�\1120130-���O�Υ��]��d\IMG_3824.JPG>2"

tmp = split(s, ",")

Dim arr

u = UBound(tmp)

ReDim arr(0 To u, 0 To 1)

For i = LBound(tmp) To UBound(tmp)

    tmp2 = split(tmp(i), ">")
    
    arr(i, 0) = tmp2(0)
    arr(i, 1) = tmp2(1)

Next

splitPhotoList = arr

End Function

Sub reCountChecks() ' not yet

With Sheets("Check")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 3 To lr
        
        origin_name = .Cells(r, 2) & "-" & .Cells(r, 3)
        
        cnt = 0
        
        For rr = 3 To r
        
            If .Cells(rr, 2) = .Cells(r, 2) Then cnt = cnt + 1
        
        Next
        
        new_name = .Cells(r, 2) & "-" & cnt
        
        If origin_name <> new_name Then
        
            .Cells(r, 3) = cnt
            
            file_path = getThisWorkbookPath & "\��d��Output\" & origin_name & ".xls"
            Call killSpecitficFile(file_path)
            
            file_path = getThisWorkbookPath & "\��d��Output\" & new_name & ".xls"
            Call killSpecitficFile(file_path)
        
        End If
    
    Next
    
    Call cmdPrintCheck

End With

End Sub

Sub killSpecitficFile(ByVal file_path As String)

Set fso = CreateObject("Scripting.FileSystemObject")

'file_path = getThisWorkbookPath & "\��d��Output\" & check_file_name & ".xls"

'If fso.fileExists(file_path) Then
    
    Set f = fso.getFile(file_path)
    Kill f

'End If

End Sub





'===================

Sub test_RngToJPG()

Call ExcelToJPGImage(Range("C1:E24"))

End Sub

Sub ExcelToJPGImage(imageRng As Range)
    'Code from officetricks.com
    Dim sImageFilePath As String
    sImageFilePath = ThisWorkbook.Path & Application.PathSeparator & "ExcelRangeToImage_"
    sImageFilePath = sImageFilePath & VBA.Format(VBA.Now, "DD_MMM_YY_HH_MM_SS_AM/PM") & ".jpg"
    
    'Create Temporary workbook to hold image
    Dim wbTemp As Workbook
    Set wbTemp = Workbooks.Add(1)
    
    'Copy image & Save to new file
    imageRng.CopyPicture xlScreen, xlPicture
    wbTemp.Activate
    With wbTemp.Worksheets("�u�@��1").ChartObjects.Add(imageRng.Left, imageRng.Top, imageRng.Width, imageRng.Height)
        .Activate
        .Chart.Paste
        .Chart.Export FileName:=sImageFilePath, FilterName:="jpg"
    End With

    'Close Temp workbook
    wbTemp.Close False
    Set wbTemp = Nothing
    'MsgBox "Image File Saved To: " & sImageFilePath
    
    frm_Photo_TMP.TextBox1.Text = sImageFilePath
    
    Call frm_Photo_TMP.Show
    
End Sub




'TODO:
'�˸���ި��`��
'�~���d��


Sub Test2()

a = "G:\�ڪ����ݵw��\ExcelVBA\�ʳy�����\@�ʳy�����DEV\�I�u�Ӥ�\1120130-���O�Υ��]��d\123.jpg"



initialFolder = mid(a, 1, InStrRev(a, "\"))

Debug.Print initialFolder

End Sub




