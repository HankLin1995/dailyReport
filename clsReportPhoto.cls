VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReportPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private shtReport As Object
Private shtResult As Object
Private wb As Object

Private coll_photo_path As New Collection
Private coll_rng_address As New Collection
Private collitem As New Collection
Private collday As New Collection
Public IsXLS As Boolean
Public IsShowText As Boolean

Private Sub Class_Initialize()

Application.ScreenUpdating = False

Set shtReport = ThisWorkbook.Sheets("ReportPhoto")
'Set shtResult = ThisWorkbook.Sheets("Result")

coll_rng_address.Add "C3"
coll_rng_address.Add "C9"
coll_rng_address.Add "C15"

End Sub


Sub GetReportByItem(ByVal r As Integer)

Dim Print_obj As New clsPrintOut

If IsXLS Then Set wb = Workbooks.Add

Call ClearAllPhoto
Call ClearAllDetail

With ThisWorkbook.Sheets("Check")

    Page = 0
    Count = 0
    photo_inf = .Cells(r, "I")
                    
    tmp = split(photo_inf, ",")
    
    tmp_canal = split(.Cells(r, "F"), ",")

    For i = LBound(tmp) To UBound(tmp)
        
        photo_inf_detail = split(tmp(i), ">")
    
        photo_path = photo_inf_detail(0)
        checkitem = .Cells(r, "B") '檢查項目
        check_num = .Cells(r, "C")
        checkch = tmp_canal(0) '.Cells(r, "G") '渠道名稱
        checkloc = tmp_canal(1) ' .Cells(r, "H") '測點
        checkday = .Cells(r, "D") '日期
        check_note = photo_inf_detail(1) '備註
        
        Count = Count + 1
        
        Call PastePhoto(Count, photo_path, checkday)
        Call PasteDetail(Count, checkitem, checkch, checkloc, checkday, check_note)
        
        If Count = 3 Then
        
            Count = 0
            Page = Page + 1
            
            If IsXLS Then
                Call PrintXLS(checkitem & "-" & check_num & "-" & Page)
            Else
                Call PrintReport(checkitem & "-" & check_num & "-" & Page)
            End If
            
            Call ClearAllPhoto
            Call ClearAllDetail

        End If
        
    Next
    
    If Count Mod 3 <> 0 Then
    
        Count = 1
        Page = Page + 1
        
        If IsXLS Then
            Call PrintXLS(checkitem & "-" & check_num & "-" & Page)
        Else
            Call PrintReport(checkitem & "-" & check_num & "-" & Page)
        End If
        
        Call ClearAllPhoto
        Call ClearAllDetail
        
    End If
    
If IsXLS Then

    Application.DisplayAlerts = False
    wb.Sheets("工作表1").Delete
    wb.SaveAs getThisWorkbookPath & "\查驗照片Output\" & checkitem & "-" & check_num & ".xls", FileFormat:=xlExcel8
    wb.Close
    Application.DisplayAlerts = True
    
End If

End With

End Sub

Sub PrintXLS(ByVal filename As String)

shtReport.Copy after:=wb.Sheets(wb.Sheets.Count)

filename = Replace(filename, "\", "1")

wb.Sheets(wb.Sheets.Count).Name = filename

End Sub

Sub PrintReport(ByVal filename As String)

wbpath = getThisWorkbookPath

shtReport.Activate

On Error Resume Next
MkDir wbpath & "\" & "查驗照片Output"
On Error GoTo 0

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:= _
wbpath & "\" & "查驗照片Output\" & filename, Quality:= _
xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
OpenAfterPublish:=False

If Err <> 0 Then
    MsgBox "你的excel版本過舊電腦無法另存為PDF!" & vbNewLine & "請點選列印紙張!": End
    Exit Sub
    
End If


End Sub

Sub CollectItem(ByVal col As String)

'"G"=查驗項目
'"J"=查驗日期

With shtResult

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        If .Cells(r, col) <> "" Then
            
            On Error Resume Next
            
            collitem.Add .Cells(r, col), CStr(.Cells(r, col))
            
        End If
    
    Next

End With

End Sub

Sub PasteDetail(ByVal Count As Integer, ByVal checkitem As String, ByVal checkch As String, _
                ByVal checkloc As String, ByVal checkday As String, ByVal check_note As String)

With shtReport

    Set rng = .Range(coll_rng_address(Count)).Offset(0, -1)
    rng.Offset(0, 0) = checkch
    rng.Offset(1, 0) = checkloc
    rng.Offset(2, 0) = tranDate(checkday) ' "'" & checkday
    rng.Offset(3, 0) = checkitem
    rng.Offset(4, -1) = check_note
    
End With

End Sub

Sub PastePhoto(ByVal Count As Integer, ByVal photo_path As String, ByVal checkday As String)

'額外接入checkday當作text基底文字

With shtReport

    Set rng = .Range(coll_rng_address(Count))

    If rng.MergeCells Then
    
        myMergeArea = rng.MergeArea
        dy = UBound(myMergeArea, 1)
        dx = UBound(myMergeArea, 2)
    
    End If
    
    Set LastRng = .Cells(rng.Row + dy - 1, rng.Column + dx - 1).Offset(1, 1)
    
    X0 = rng.Left
    Y0 = rng.Top
    
    X1 = LastRng.Left
    Y1 = LastRng.Top
    
    Set pic = .Shapes.AddPicture(photo_path, True, True, X0 + 2, Y0 + 2, X1 - X0 - 4, Y1 - Y0 - 4)
    pic.LockAspectRatio = msoFalse
    
    If IsShowText = True Then Call AddText(X1 - 100, Y1 - 30, 25, 40, tranDate(checkday), 3) '額外接入之checkday要顯示於照片右下角
    
End With

End Sub

Function tranDate(ByVal dateStr As String)

a = mid(dateStr, 1, 4)
b = mid(dateStr, 5, 2)
c = mid(dateStr, 7, 2)

mydate = a & "/" & b & "/" & c

tranDate = mydate

End Function

Sub AddText(ByVal txtX, ByVal txtY, ByVal txtHeight, ByVal txtWidth, ByVal txtStr, Optional ByVal RGB_Selector As Integer = 0)

'If ThisWorkbook.Sheets("Main").Range("B4") = "N" Then Exit Sub

If txtStr = specific_date Then RGB_Selector = 3

With shtReport

    .Activate

    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, txtX, txtY, txtHeight * Len(txtStr) * 0.33, txtHeight).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = txtStr
        
    Selection.ShapeRange.Fill.Visible = msoFalse
    Selection.ShapeRange.Line.Visible = msoFalse
    Selection.ShapeRange.ZOrder msoBringToFront
        ' Selection.ShapeRange.TextFrame2.TextRange.Font.Bold = msoTrue
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Font
'        .NameComplexScript = "+mn-cs"
'        .NameFarEast = "+mn-ea"
'        .Fill.Visible = msoTrue
'        .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
'        .Fill.ForeColor.TintAndShade = 0
'        .Fill.ForeColor.Brightness = 0
'        .Fill.Transparency = 0
'        .Fill.Solid
        .Bold = msoTrue
        Select Case RGB_Selector
        
        Case 0: .Fill.ForeColor.RGB = RGB(0, 0, 0)
        Case 1: .Fill.ForeColor.RGB = RGB(0, 204, 255)
        Case 2: .Fill.ForeColor.RGB = RGB(255, 0, 0)
        Case 3: .Fill.ForeColor.RGB = RGB(255, 102, 0)
        
        End Select
        
        '.Fill.Visible = msoFalse
        .Size = txtHeight * 0.5
'        .Name = "+mn-lt"
    End With

End With

End Sub

Sub ClearAllPhoto() '清除Report舊照片

With shtReport

For Each shp In .Shapes
    
    shp.Delete
    
Next

End With

End Sub

Sub ClearAllDetail()

With shtReport
    
    arr = Array("B3", "B9", "B15")
    
    For Each ad In arr
    
        .Range(ad).Resize(4, 1).ClearContents
    
    Next
    
    .Range("A7") = ""
    .Range("A13") = ""
    .Range("A19") = ""

End With

End Sub

Private Sub Class_Terminate()

'If IsXLS = True Then
    
    'Application.DisplayAlerts = False
    'wb.Sheets(1).Delete
    'Application.DisplayAlerts = True
    
'End If

Application.ScreenUpdating = True

End Sub


