VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ReportDay As Date
Private ItemArr As Variant
Private t_change As Integer

'Private SigLeft As Single
'Private SigTop As Single
'Private SigWidth As Single
'Private SigHeight As Single

Private shtReport As Object
Private shtMain As Object
Private shtDiary As Object
Private shtRecords As Object

Private myfunc As New clsMyfunction

Private Sub Class_Initialize()

Set shtReport = Sheets("Report")
Set shtDiary = Sheets("Diary")
Set shtRecords = Sheets("Records")
Set shtMain = Sheets("Main")

'On Error Resume Next

ReportDay = shtReport.Range("C2")

'checkIfInWorkPeriod

ReportNum = shtReport.Range("K2")

key_value = ReportNum
Set rng = shtDiary.Columns("A").Find(what:=key_value, LookIn:=xlValues)

If rng Is Nothing Then MsgBox "編號在Diary中找不到!", vbCritical: End

'If Err <> 0 Then shtReport.Range("K2") = 1: End

End Sub

Sub ResetReport()

With shtReport

'For r = 1 To 100
'
'    If .Cells(r, 1) Like "*施工項目*" Then
'        lr = getReportLastRow - 1
'        Exit For
'    End If
'
'Next

lr = getReportLastRow - 1

Do Until i = lr - 8 - 4
.Rows("10").Delete
i = i + 1
Loop

End With

End Sub

Function getReportLastRow() '2022-11-23新增

With Sheets("Report")

    Set rng = .UsedRange.Find("一、工程進行情況（含約定之重要施工項目及數量）：")
    
    If rng Is Nothing Then
        lr = InputBox("找不到【一、工程進行情況（含約定之重要施工項目及數量）：】" & vbNewLine & "請輸入此行應於第幾行?")
    Else
        lr = rng.Row
    End If

    getReportLastRow = lr

End With

End Function

Sub CollectItem() '取得報表資料

c = 6 + t_change * 5

With shtMain

    If .Cells(1, c) = "" Then
        MsgBox "查無該日有報表資料!請先進行變更設計資料有無匯入至報表!", vbCritical: Call cmdExportToReport
    End If

    lr = .Cells(Rows.Count, c).End(xlUp).Row
    
    ItemCount = lr - 3 + 1
    
    ReDim ItemArr(1 To ItemCount, 1 To 7)
    
    j = 1
    
    For r = 3 To lr

        ItemName = .Cells(r, c)
        ItemUnit = .Cells(r, c + 1)
        ItemContractNum = .Cells(r, c + 2)
        ItemMoney = .Cells(r, c + 3)
        ItemSep = .Cells(r, c).Font.ColorIndex
        
        arr = Array(ItemName, ItemUnit, ItemContractNum, ItemMoney)
        
        For i = 0 To UBound(arr)
        
            ItemArr(j, i + 1) = arr(i)
            
        Next
        
        ItemArr(j, 7) = ItemSep
        
        j = j + 1
        
    Next

End With

End Sub

Sub CollectRec() '收集本日數量與累計數量

Dim REC_obj As New clsRecord

With shtRecords

    lr = .Cells(Rows.Count, 1).End(xlUp).Row

    For i = 1 To UBound(ItemArr, 1)
    
        ItemSep = ItemArr(i, 7)
    
        TargetItemName = ItemArr(i, 1)
        
        Call REC_obj.getNumAndSumByItemName(TargetItemName, ReportDay, RecAmount, RecSumAmount)
        
        ItemArr(i, 5) = RecAmount
        ItemArr(i, 6) = RecSumAmount 'WorksheetFunction.RoundDown(RecSumAmount, 0)
    
    Next

End With

Call GetPercentItem

End Sub

Private Sub GetPercentItem()

For i = 1 To UBound(ItemArr, 1) 'only for colorindex=3,當分母

    ItemContractNum = ItemArr(i, 3)
    ItemMoney = ItemArr(i, 4)
    ItemSep = ItemArr(i, 7)

    If ItemSep = 3 Then totalmoney = totalmoney + ItemContractNum * ItemMoney

Next

For i = 1 To UBound(ItemArr, 1)

    ItemSep = ItemArr(i, 7)
    
    If ItemSep = 3 Then
    
        ItemContractNum = ItemArr(i, 3)
        ItemMoney = ItemArr(i, 4)
        ItemAmount = ItemArr(i, 5)
        ItemSumAmount = ItemArr(i, 6)
        
        ReportMoney_P = ReportMoney_P + ItemAmount * ItemMoney
        SumReportMoney_P = SumReportMoney_P + ItemSumAmount * ItemMoney

        arr = Array(ReportMoney_P / totalmoney, SumReportMoney_P / totalmoney)

    End If

Next

For i = 1 To UBound(ItemArr, 1)

    ItemSep = ItemArr(i, 7)

    If ItemSep = 5 Then
    
        ItemArr(i, 5) = arr(0)
        ItemArr(i, 6) = arr(1)

    End If

Next

End Sub

Private Sub checkItemsRow()

With shtReport

    '------------格式重整----------

    lr = getReportLastRow - 1

    .Cells(8, 2).Resize(lr - 8 + 1, 8).ClearContents

    i = 1

    Itemlr = UBound(ItemArr, 1) + 8 - 1

    If Itemlr >= lr Then '直接新增欄位
    
        For i = 1 To Itemlr - lr + 1
        
            Rows("9:9").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("B8:K8").AutoFill Destination:=Range("B8:K9"), Type:=xlFillDefault
        
            lr = lr + 1
        
        Next
        
        Dim obj3 As New clsReportTest
        
        obj3.Init_Test
    
        MsgBox "已新增施作項目欄位!試驗位置已經重整!!"
        
    End If
    
End With

End Sub

Sub WriteReport() '***

Dim pgs_today As Double
Dim pgs_total As Double

Call checkItemsRow

'Call GetPercentItem

With shtReport

'    '------------格式重整----------
'
    lr = getReportLastRow - 1
'
'    .Cells(8, 2).Resize(lr - 8 + 1, 8).ClearContents
'
'    i = 1
'
'    Itemlr = UBound(ItemArr, 1) + 8 - 1
'
'    If Itemlr >= lr Then '直接新增欄位
'
'        For i = 1 To Itemlr - lr + 1
'
'            Rows("9:9").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
'            Range("B8:K8").AutoFill Destination:=Range("B8:K9"), Type:=xlFillDefault
'
'            lr = lr + 1
'
'        Next
'
'        Dim obj3 As New clsReportTest
'
'        obj3.Init_Test
'
'        MsgBox "已新增施作項目欄位!試驗位置已經重整!!"
'
'    End If
'
    '-------------------------------
    
    
    i = 1
    
    For r = 8 To UBound(ItemArr, 1) + 8 - 1
    
        ItemName = ItemArr(i, 1)
        ItemUnit = ItemArr(i, 2)
        ItemContractNum = ItemArr(i, 3)
        ItemMoney = ItemArr(i, 4)
        ItemAmount = ItemArr(i, 5)
        ItemSumAmount = ItemArr(i, 6)
        ItemSep = ItemArr(i, 7)
        
        .Range("B" & r) = ItemName
        .Range("E" & r) = ItemUnit
        .Range("F" & r) = ItemContractNum
        .Range("G" & r) = ItemAmount
        .Range("I" & r) = ItemSumAmount
    
        totalmoney = totalmoney + ItemContractNum * ItemMoney
        ReportMoney = ReportMoney + ItemAmount * ItemMoney
        SumReportMoney = SumReportMoney + ItemSumAmount * ItemMoney
        
        .Range("I" & r).Interior.ColorIndex = -4142
        
        If Int(ItemSumAmount) > Int(ItemContractNum) Then
            .Range("I" & r).Interior.ColorIndex = 44
            'Stop
        End If
        
        i = i + 1
        
        
    Next
    
    .Range("B" & lr) = "發包工作費總額"
    .Range("F" & lr) = totalmoney
    .Range("G" & lr) = WorksheetFunction.RoundDown(ReportMoney, 0)
    .Range("I" & lr) = WorksheetFunction.RoundDown(SumReportMoney, 0)
    
    '-----進度部分-----
    
    pgs_today = .Range("G" & lr) / .Range("F" & lr)
    pgs_total = .Range("I" & lr) / .Range("F" & lr)


    '先無條件捨棄後放入Diary
    
    Call KeyInPGS(.Range("C2").Value, pgs_total)
    
    '取得正確的pgs_todya及pgs_total
    
    Call getCorrectPgs(.Range("C2").Value, pgs_today, pgs_total)

    .Range("F6") = pgs_today
    .Range("I6") = pgs_total

End With

'Dim obj2 As New clsReportTest
'
'obj2.t_change = t_change
'obj2.CollTestArr
'obj2.GetTestRec
'obj2.WriteReport

End Sub

Sub getCorrectPgs(ByVal mydate As Date, ByRef pgs_today As Double, ByRef pgs_total As Double)

With Sheets("Diary")
    
    Set coll_rows = myfunc.getRowsByUser2("Diary", Format(mydate, "yyyy/mm/dd(aaa)"), 1, "報表日期")
    
    If coll_rows.Count <> 0 Then
        
        r = coll_rows(1)
        
        If r = 2 Then
        
            pgs_today = .Cells(r, "I")
            pgs_total = pgs_today
        
        Else
        
            pgs_total_1 = .Cells(r - 1, "I")
            pgs_total_2 = .Cells(r, "I")
            
            pgs_today = pgs_total_2 - pgs_total_1
            pgs_total = pgs_total_2
    
        End If
    
    End If

End With


End Sub

Sub WriteReport_Test()

Dim obj2 As New clsReportTest

obj2.t_change = t_change
obj2.CollTestArr
obj2.GetTestRec
obj2.WriteReport

End Sub

Sub KeyInPGS(ByVal mydate As Date, ByVal pgs_total As Double)

Set coll_rows = myfunc.getRowsByUser2("Diary", Format(mydate, "yyyy/mm/dd(aaa)"), 1, "報表日期")
If coll_rows.Count <> 0 Then Sheets("Diary").Cells(coll_rows(1), 9) = WorksheetFunction.Round(pgs_total, 4)

End Sub

Sub getInfo() '取得Main工作表資訊

Dim PCCES_obj As New clsPCCES
Dim Inf_obj As New clsInformation

'Inf_obj.getBasicInformation


'arr = obj.ReturnInfoArr
'
'For i = 1 To UBound(arr)
'
'    Title = arr(i, 1)
'
'    'If Title Like "coll*" Then
'        'Set Data = arr(i, 2)
'    'Else
'    Data = arr(i, 2)
'    'End If
'
'    Select Case Title
'
'    'Case "IsChanged": IsChanged = Data
'    Case "IsEnlarged": IsEnlarged = Data
'    Case "EnlargedDate": EnlargedDate = Data
'    Case "WorkDay": workDay = Data
'    Case "ContractMoney": contractMoney = Data
'    'Case "collContractMoney": Set collChangedMoney = Data
'    'Case "collContractDay": Set collChangedDay = Data
'    Case "WorkDayAdd": workDayAdd = Data
'
'    End Select
'
'Next

IsEnlarged = Inf_obj.IsEnlarged
EnlargedDate = Inf_obj.EnlargedDate
workDay = Inf_obj.workDay
workDayAdd = Inf_obj.workDayAdd
endDate = Inf_obj.GetEndDate

With shtReport

    ReportDay = .Range("C2")
    
    '-----展延工期相關------
    
    If ReportDay < EnlargedDate Then workDay = workDay - workDayAdd
    
    .Range("C4") = workDay
    .Range("I4") = endDate
    
    If IsEnlarged = True And ReportDay >= EnlargedDate Then
        .Range("H5") = Sheets("Main").Cells(6, 4) & workDayAdd & "天"
    Else
        .Range("H5") = "天"
    End If
    
    
    '-----------------------

    .Range("K5") = "原契約:" & PCCES_obj.getSumMoney
    .Range("K6") = "變更後契約:"
    .Range("D5") = "次"
    
    t_change = Inf_obj.getContractChangesByDate(ReportDay)
    
    If t_change > 0 Then
    
    .Range("K6") = "變更後契約:" & PCCES_obj.getSumMoney(t_change)
    .Range("D5") = t_change & " 次"
    
    End If
    
'    If collChangedMoney.count > 1 Then
'
'        For i = collChangedMoney.count To 2 Step -1
'
'            If ReportDay >= collChangedDay.Item(i) Then
'                .Range("K6") = "變更後契約:" & collChangedMoney.Item(i)
'                t_change = t_change + 1
'                .Range("D5") = t_change & " 次"
'                Exit For
'            End If
'        Next
'
'    End If

End With

End Sub
'
'Private Sub GetMainRow()
'
'Call ReturnMainRow(arr)
'
'ConRow = arr(0)
'TestRow = arr(1)
'ChannelRow = arr(2)
'ItemRow = arr(3)
'PgsRow = arr(4)
'
'End Sub

'Sub CreateSig()
'
'With shtReport
'
'    SigLeft = .Range("J8")
'    SigTop = .Range("J9")
'    SigWidth = .Range("J10")
'    SigHeight = .Range("J11")
'
'End With
'
'Randomize
'
'rand = Int(Rnd() * 6)
'
'Path = getThisWorkbookPath & "\簽名" & rand + 1 & ".JPG"
'
'Set pics = ActiveSheet.Pictures
'
'pics.Delete
'
'pics.Insert (Path)
'
'Set pic = pics.Item(1)
'
'pic.Left = SigLeft + Rnd() * 20
'pic.Top = SigTop + Rnd()
'pic.Width = SigWidth - Rnd() * 2
'pic.Height = SigHeight - Rnd() * 2
'
'End Sub


'Sub hideRow()
'
'With Sheets("Report")
'
'.Rows("8:76").Hidden = False
'
'For r = 8 To 76
'
'    '.Rows(r).Hidden = False
'
'    If .Cells(r, "I") = 0 Then .Rows(r).Hidden = True 'Debug.Print .Cells(r, 2)
'
'Next
'
'For r = 59 To 76
'
'    If r <> 75 Then .Rows(r).Hidden = True
'
'Next
'
'End With

'End Sub



