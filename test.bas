Attribute VB_Name = "test"

Sub test_collectFilePaths()

Dim f As New clsMyfunction
Dim coll As New Collection

With Sheets("Check")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 3 To lr
    
        MyCode = .Cells(r, 2)
        myNum = .Cells(r, 3)
    
        myPath = ThisWorkbook.Path & "\抽查表PDF\" & MyCode & "-" & myNum & ".pdf"
    
        If f.IsFileExists(myPath) = True Then
            
            coll.Add myPath
        
        End If
        
        myPath = ThisWorkbook.Path & "\查驗照片Output\" & MyCode & "-" & myNum & ".pdf"
    
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

    ' 設置可執行文件所在的目錄
    'path1 = ThisWorkbook.Path ' ThisWorkbook.Path & "\GooglePhotoDownloader"
    'ChDir path1

    PDF_PATH = ThisWorkbook.Path & "\Lib\Merge\merge.pdf"
    TXT_PATH = ThisWorkbook.Path & "\Lib\Merge\file_with_paths.txt"
    
    command = ThisWorkbook.Path & "\Lib\Merge\Merge.exe """ & TXT_PATH & """ """ & PDF_PATH & """"
    
    ' 使用Shell函數運行可執行文件並等待其完成
    myret = Shell(command, 1)

End Sub

Sub CheckFileCreateTimeWithRetryAndDelay(ByVal CurrentTime)
    ' 文件路?
    FilePath = ThisWorkbook.Path & "\Lib\merge\merge.pdf"
    
    ' ?置容忍的差异??（?里?置?1小?）
    ToleranceHours = 0 '1 / 24
    
    ' 最大??次?
    MaxTries = 20
    
    Dim FileCreateTime As Date
    'Dim CurrentTime As Date
    Dim Tries As Integer
    
    Tries = 1
    
    Do While Tries <= MaxTries
    
        Debug.Print Tries
    
        If Dir(FilePath) = "" Then
            ' 文件不存在，等待2秒?再??
            Tries = Tries + 1
            If Tries <= MaxTries Then
                Application.Wait Now + TimeValue("00:00:02")
            End If
        Else
            ' 文件存在，?取文件的?建??
            FileCreateTime = FileDateTime(FilePath)
            
            ' ?取?前??
            'CurrentTime = Now
            
            ' 比?文件?建??和?前??，考?容忍??
            If FileCreateTime > CurrentTime Then
                Shell "cmd /c start " & FilePath, vbNormalFocus
                Exit Do
            Else
                ' 文件?建??在容忍范?外，等待2秒?然后重?
                Tries = Tries + 1
                If Tries <= MaxTries Then
                    Application.Wait Now + TimeValue("00:00:02")
                End If
            End If
        End If
    Loop
    
    If Tries > MaxTries Then
        MsgBox "已?到最大重?次?。"
    End If
End Sub




Sub WriteCollectionToTxt(ByVal coll)
    'Dim col As Collection
    Dim Item As Variant
    Dim FilePath As String
    Dim FileName As String
    Dim FileNumber As Integer
    
    ' ?建一?新的集合并添加?据（?里只是示例，你需要自己填充你的集合）
    'Set col = New Collection
    'col.Add "Item1"
    'col.Add "Item2"
    'col.Add "Item3"
    
    ' 指定文件名和路?
    FileName = "file_with_paths.txt"
    FilePath = ThisWorkbook.Path & "\Lib\Merge\" & FileName
    
    ' 打?文本文件以?行?入
    FileNumber = FreeFile
    Open FilePath For Output As FileNumber
    
    ' 遍?集合并?每??目?入文本文件
    For Each Item In coll
        Print #FileNumber, Item
    Next Item
    
    ' ??文件
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
    
    If newFileName = "" Then MsgBox "請先框選要歸檔的位置!", vbCritical

    ' 讓使用者選擇一個檔案
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "選擇要複製的檔案"
        If .Show = -1 Then
            SourceFilePath = .SelectedItems(1)
        Else
            'MsgBox "未選擇檔案。操作已取消。"
            Exit Sub
        End If
    End With

    ' 指定目標資料夾
    DestinationFolder = ThisWorkbook.Path & "\抽查表PDF\"

    ' 獲得原始檔案名稱
    SourceFileName = mid(SourceFilePath, InStrRev(SourceFilePath, "\") + 1)

    ' 組合新的檔案路徑
    NewFilePath = DestinationFolder & newFileName & ".pdf" 'SourceFileName

    ' 檢查目標資料夾是否存在，如不存在則創建
    If Dir(DestinationFolder, vbDirectory) = "" Then
        MkDir DestinationFolder
    End If

    ' 複製檔案
    FileCopy SourceFilePath, NewFilePath

    ' 檢查檔案是否成功複製
    
    With Sheets("Check")
    
        If Dir(NewFilePath) <> "" Then
            'MsgBox "檔案已成功複製到 " & NewFilePath
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
        
        If pcces_item Like "*試驗規範及標準*" Then
            
            If pcces_item Like "*;*" Then MsgBox pcces_item & vbNewLine & "含有非法字符【;】，請修正!", vbCritical: End
            
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

        getTestReplaceName = InputBox(test_item & ":未定義別名!" & vbNewLine & "請輸入別名:", , test_item)
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

If coll.Count = 0 Then MsgBox "請先開啟要匯入的舊檔!", vbCritical: End

For i = 1 To coll.Count

    p = p & i & "." & coll(i) & vbNewLine

Next

If coll.Count = 1 Then

getWorkbookName = coll(1)

Else

myIndex = InputBox("請輸入要匯入的檔案名稱" & vbNewLine & p)

getWorkbookName = coll(myIndex)

End If

End Function































Function splitPhotoList(ByVal s As String)

If s = "" Then Exit Function

s = "G:\我的雲端硬碟\ExcelVBA\監造日報表\@監造日報表DEV\上課素材\施工照片\1120130-鋼板樁打設抽查\IMG_3824.JPG>1,G:\我的雲端硬碟\ExcelVBA\監造日報表\@監造日報表DEV\上課素材\施工照片\1120130-鋼板樁打設抽查\IMG_3824.JPG>2"

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
            
            file_path = getThisWorkbookPath & "\抽查表Output\" & origin_name & ".xls"
            Call killSpecitficFile(file_path)
            
            file_path = getThisWorkbookPath & "\抽查表Output\" & new_name & ".xls"
            Call killSpecitficFile(file_path)
        
        End If
    
    Next
    
    Call cmdPrintCheck

End With

End Sub

Sub killSpecitficFile(ByVal file_path As String)

Set fso = CreateObject("Scripting.FileSystemObject")

'file_path = getThisWorkbookPath & "\抽查表Output\" & check_file_name & ".xls"

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
    With wbTemp.Worksheets("工作表1").ChartObjects.Add(imageRng.Left, imageRng.Top, imageRng.Width, imageRng.Height)
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
'檢試驗管制總表
'品質抽查表


Sub Test2()

a = "G:\我的雲端硬碟\ExcelVBA\監造日報表\@監造日報表DEV\施工照片\1120130-鋼板樁打設抽查\123.jpg"



initialFolder = mid(a, 1, InStrRev(a, "\"))

Debug.Print initialFolder

End Sub




