Attribute VB_Name = "Normal"

Sub t()

Dim o As New clsPrintOut

o.ToXLS_test

End Sub

Sub ToXLS()

Application.DisplayAlerts = False

sr = Val(InputBox("開始頁數"))
er = Val(InputBox("結束頁數"))

For r = er To sr Step -1

    Debug.Print "列印頁面=" & r

    ThisWorkbook.Activate

    ThisWorkbook.Sheets("Report").Range("K2") = r
    
    Call ReportRun

Next

Application.DisplayAlerts = True

End Sub


Sub changeNum()

myIndex = InputBox("請輸入編號")

If myIndex = "" Then Exit Sub

ActiveSheet.Range("K2") = myIndex

Call ReportRun

End Sub


Sub ReportRun()

Dim obj As New clsReport

'obj.CreateSig

obj.getInfo
obj.CollectItem
obj.CollectRec
obj.WriteReport
obj.WriteReport_Test
'obj.hideRow

End Sub

Sub ResetItem()

On Error Resume Next

If Not frmData.txtDay Like "*/*/*" Then Exit Sub

If frmData.cboItem = "" Then Exit Sub

frmData.txtAmount = 0
'frmData.cboItem = ""

Dim obj As New clsBasicData

Call obj.RetrunUnit
Call obj.UsedAmount

'Call checkTestCompleted

End Sub

Sub ResetItem_Mix()

On Error Resume Next

If Not MixData_Main.txtDay Like "*/*/*" Then Exit Sub

If MixData_Main.cboItem = "" Then Exit Sub

MixData_Main.txtAmount = 0

Dim obj As New clsMixData

 obj.ReadData
obj.ReturnLast
obj.UsedAmount

End Sub


Sub ReturnMainRow(arr)

ReDim arr(5)

With Sheets("Main")
    
    tmp = Split(.Cells(2, 3).Value, ",")
    
    For i = 0 To 5
        On Error Resume Next
        arr(i) = tmp(i)
    
    Next
    
    ResetDay = arr(5)
    
    'If ResetDay = Format(Now, "yyyy/mm/dd") Then Exit Sub
    
    For Each rng In .UsedRange
    
        Select Case rng.Value
        
        Case "工程名稱": arr(0) = rng.Row
        Case "試體名稱": arr(1) = rng.Row
        Case "施工渠道名稱": arr(2) = rng.Row
        Case "工程項目": arr(3) = rng.Row
        Case "累積進度(%)": arr(4) = rng.Row
        
        End Select
    
    Next
    
    .Cells(2, 3) = arr(0) & "," & arr(1) & "," & arr(2) & "," & arr(3) & "," & arr(4) & "," & Format(Now, "yyyy/mm/dd")

End With

End Sub

