Attribute VB_Name = "test"
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




