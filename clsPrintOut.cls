VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrintOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Filename As String
Private sht As Object

Sub BeforePrintCheck()

'For Each rng In Sheets("Report").UsedRange

    'If rng.Interior.ColorIndex = 44 Then MsgBox "累計數量超過契約數量囉~": End

'Next

End Sub

Sub ToPDF() '給包商參考用的

sr = Val(InputBox("開始頁數"))

er = Val(InputBox("結束頁數"))

wbpath = ThisWorkbook.Path

For i = sr To er

    Sheets("Report").Activate
    
    Range("K2") = i
    
    Call ReportRun
    
        On Error Resume Next
        MkDir wbpath & "\" & "監造日報表Output"
        On Error GoTo 0
        'ChDir wbpath
        
        On Error Resume Next
        
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        wbpath & "\" & "監造日報表Output\" & Range("K2"), Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
        
        If Err <> 0 Then
            MsgBox "你的excel版本過舊電腦無法另存為PDF!" & vbNewLine & "請點選列印紙張!": End
            Exit Sub
            
        End If
            
        On Error GoTo 0
        
Next

Shell "explorer.exe " & wbpath & "\" & "監造日報表Output\", vbNormalFocus

End Sub

Sub ToPaper()

msg = MsgBox("目前的印表機為:" & Application.ActivePrinter & vbNewLine & "是否要繼續列印?", vbYesNo + vbInformation)

If msg = vbNo Then End

MsgBox "這個按鈕太危險~先印成電子檔吧!"

Exit Sub

sr = Val(InputBox("開始頁數"))
er = Val(InputBox("結束頁數"))

wbpath = ThisWorkbook.Path

For i = sr To er

    Sheets("Report").Activate
    
    Range("K2") = i
        
Next

End Sub

Sub ToPDF_Check(ByVal sht As Worksheet, ByVal fname As String, ByVal mode) '查驗表輸出成PDF

    wbpath = ThisWorkbook.Path
    
    If mode = 1 Then
    
    On Error Resume Next
    MkDir wbpath & "\" & "查驗表Output"
    On Error GoTo 0
    
    On Error Resume Next
    
    sht.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
    wbpath & "\" & "查驗表Output\" & fname, Quality:= _
    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
    OpenAfterPublish:=False

    If Err <> 0 Then MsgBox "你的excel版本過舊電腦無法另存為PDF!" & vbNewLine & "請點選列印紙張!": End
    
    Else
    
    sht.PrintOut
    Exit Sub
    End If
    
    'Shell "explorer.exe " & wbpath & "\" & "查驗表Output\", vbNormalFocus
    
End Sub

'ToWorksheets

Sub ToXLS_test()

Application.DisplayAlerts = False

sr = Val(InputBox("開始頁數"))
er = Val(InputBox("結束頁數"))

Set wb = Workbooks.Add

For r = er To sr Step -1

    Debug.Print "列印頁面=" & r

    ThisWorkbook.Activate

    ThisWorkbook.Sheets("Report").Range("K2") = r
    
    Dim obj As New clsReport
    
    obj.WriteReport_Test
    obj.hideRow
    
    ThisWorkbook.Sheets("Report").Copy (wb.Sheets(1))
    
    Set sht = wb.Sheets(1)
    
    For Each shp In sht.Shapes
    
        shp.Delete
        
    Next
    
    With sht
        
        .Name = r
        .Columns("A:K").Select
        .Range("A7").Activate
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    End With
    
Next

wb.Sheets("工作表1").Delete

Application.DisplayAlerts = True

End Sub

Sub ToXLS()

Application.DisplayAlerts = False

sr = Val(InputBox("開始頁數"))
er = Val(InputBox("結束頁數"))

Set wb = Workbooks.Add

For r = er To sr Step -1

    Debug.Print "列印頁面=" & r

    ThisWorkbook.Activate

    ThisWorkbook.Sheets("Report").Range("K2") = r
    
    Call ReportRun
    
    ThisWorkbook.Sheets("Report").Copy (wb.Sheets(1))
    
    Set sht = wb.Sheets(1)
    
    For Each shp In sht.Shapes
    
        shp.Delete
        
    Next
    
    With sht
        
        .Name = r
        .Columns("A:K").Select
        .Range("A7").Activate
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
    End With
    
Next

wb.Sheets("工作表1").Delete

Application.DisplayAlerts = True

End Sub

Sub SpecificShtToXLS(ByVal shtName As String)

Application.DisplayAlerts = False

Set wb = Workbooks.Add
    
    ThisWorkbook.Sheets(shtName).Copy (wb.Sheets(1))
    
    Set sht = wb.Sheets(1)
    
    With sht
        
        .Name = shtName
    
    End With

wb.Sheets("工作表1").Delete

Application.DisplayAlerts = True

End Sub



