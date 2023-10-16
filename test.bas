Attribute VB_Name = "test"
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
        .Chart.Export filename:=sImageFilePath, FilterName:="jpg"
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




