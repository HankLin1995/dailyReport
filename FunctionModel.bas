Attribute VB_Name = "FunctionModel"
Sub cmdGetReportIDByDate(Optional myNewDate As Date) '20221125依日期選擇頁數

With Sheets("Report")

    mydate = .Range("C2")
    myID = .Range("K2")
    On Error GoTo DATEFORMATERRORHANDLE
    If myNewDate = 0 Then myNewDate = InputBox("請輸入日期，格式如" & vbNewLine & mydate, , mydate)
    'On Error GoTo DATEFORMATERRORHANDLE
    myNewID = myID + CDate(myNewDate) - mydate
    Set rng = Sheets("Diary").Columns("A").Find(myNewID)
    
    myDiaryDate = rng.Offset(0, 1)

    If myDiaryDate = CDate(myNewDate) Then
    
        .Range("K2") = myNewID
        Call ReportRun
    
    Else
    
        MsgBox "Diary日期不連續，請進行切換頁數!", vbCritical
    
    End If


End With

Exit Sub

DATEFORMATERRORHANDLE: MsgBox "日期格式有誤，請依照正確格式!", vbCritical

End Sub


Sub getOverNumberFromLastDay() '20221122處理剩餘零星數量

Dim obj As New clsReport

ReportNum = InputBox("請輸入理應為100%的報表編號")
allowence = InputBox("請輸入校正回歸允許值", , 1)
prompt = "***校正回歸完成項目***" & vbNewLine

With Sheets("Report")

    .Range("K2") = ReportNum

    Call ReportRun
    
    For r = 8 To obj.getReportLastRow
    
        'If r = 10 Then Stop
    
        conNum = .Cells(r, "F")
        sumNum = .Cells(r, "I")
        
        If conNum <> sumNum And conNum <> 1 Then
        
            ItemName = .Cells(r, "B")
            numDiff = Round(sumNum - conNum, 4)
            
            If Abs(numDiff) < CDbl(allowence) And numDiff <> 0 Then
            
                Call dealOverNum(ItemName, numDiff)
            
                prompt = prompt & vbNewLine & ItemName & ":" & numDiff
        
            End If
        
        End If
    
    Next

    If Len(prompt) > 16 Then
        MsgBox prompt, vbInformation
    Else
        MsgBox "查無需要校正回歸項目!", vbInformation
    End If
End With

End Sub

Sub dealOverNum(ByVal ItemName As String, ByVal numDiff As Double) '20221122處理剩餘零星數量

With Sheets("Records")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = lr To 3 Step -1
    
        recName = .Cells(r, "E")
        
        If recName = ItemName Then
        
            originNum = .Cells(r, "F")
            
            adjustNum = originNum - numDiff
            
            If adjustNum > 0 Then
            
                Debug.Print ItemName & ",原數量=" & originNum & ">>校正=" & adjustNum
            
                On Error Resume Next
            
                .Cells(r, "F").Comment.Delete
                
                On Error GoTo 0
            
                .Cells(r, "F").AddComment "originNum=" & .Cells(r, "F") & ">>adjustNum=" & adjustNum
            
                .Cells(r, "F") = adjustNum
                .Cells(r, "F").Font.ColorIndex = 7
                
                Exit For
            
            End If
        
        End If
        
    Next

End With

End Sub


Sub cmdGetReviseMixItem()

Dim o As New clsRecord

With Sheets("Mix")

lr = .Cells(.Rows.Count, "D").End(xlUp).Row

For r = 3 To lr

    If .Cells(r, 1) <> "" Then
    
        mix_name = .Cells(r, 1)
        
        Call o.ChangeMixToRecord(mix_name)
        
    End If

Next

End With

'Sheets("Records").Activate

End Sub

Sub test_getMixSumUnit()

mylen = InputBox("單元總長=?")
myName = InputBox("單元名稱")

Dim coll As New Collection

With Sheets("Mix_Sum")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 3 To 86
    
        If .Rows(r).Hidden = False Then
        
            Item = .Cells(r, 2)
            On Error Resume Next
            coll.Add Item, Item
            On Error GoTo 0
            
        End If
        
    Next

rr = 3

For Each it In coll

    Sum = 0

    For r = 3 To 86
    
        If .Rows(r).Hidden = False Then
            
            Item = .Cells(r, 2)
            
            If Item = it Then
            
                num = .Cells(r, 3) / mylen
                
                Sum = Sum + num
            
            End If
            
        End If
    
    Next
    
    With Sheets("Mix_Sum_UNIT")
    
        lr = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        
        .Cells(lr, 1) = it
        .Cells(lr, 2) = WorksheetFunction.Round(Sum, 3)
    
    End With

    '.Range("K" & rr) = it
    '.Range("L" & rr) = Sum

     s = s & vbNewLine & it & ":" & Sum

    rr = rr + 1

Next

End With

MsgBox s

End Sub

Sub cmdGetTmpData()

Dim obj As New clsRecord

'msg = MsgBox("是否要留白?", vbYesNo)

'If msg = vbYes Then
Call obj.getTmpData 'True)
'Else
'obj.getTmpData (False)
'End If

'obj.Tmp2TmpTotal

End Sub

Sub cmdMixComplete()

Dim obj As New clsMixData

obj.CheckComplete
obj.CheckUnfoundMixName

End Sub

Sub cmdResetReport()

Dim obj As New clsReport

obj.ResetReport

Sheets("Report").Activate

Call ReportRun

Sheets("Main").Activate

End Sub

Sub cmdOutput_Paper()

Dim obj As New clsPrintOut

obj.BeforePrintCheck
obj.ToPaper

End Sub

Sub cmdOpenCheck()

Dim coll_rows As New Collection

For Each rng In Selection

    On Error Resume Next
    coll_rows.Add rng.Row, CStr(rng.Row)
    On Error GoTo 0

Next

With Sheets("Check")

For Each r In coll_rows

    MyCode = .Cells(r, 2)
    myNum = .Cells(r, 3)
    
    If MyCode <> "" Then
        
        myFilePath = getThisWorkbookPath & "\抽查表Output\" & MyCode & "-" & myNum & ".xls"
    
        Workbooks.Open (myFilePath)

    End If

Next

End With

End Sub

Sub cmdPrintCheck()
'
'MsgBox "測試中..."
'
'Exit Sub

Dim obj As New clsCheck

obj.CheckList
obj.printCheckTable

End Sub

Sub cmdExportToCheck()

MsgBox "還在修理中!", vbCritical: Exit Sub

Dim obj As New clsCheck

obj.ExportToCheck
obj.CountCheck
obj.CheckList

'MsgBox "傳送完畢!!!請至""Check""查看"

Sheets("Check").Activate

End Sub

Sub cmdExportToReport()

'Dim obj As New clsBudget
'
'obj.CollectTitle
'obj.clearOldReport
'obj.ExportToReport 'should change something

Dim o As New clsInformation

Set coll = o.getContractChanges

For Each it In coll
    cnt = cnt + 1
    prompt = prompt & cnt - 1 & "." & it & vbNewLine

Next

t_change = InputBox("請輸入欲匯出至Main的編號" & vbNewLine & prompt, , cnt - 1)

With Sheets("Main")

    c = 6 + (t_change - 1) * 5
    
    If t_change > 0 Then
    
        .Cells(1, c).Resize(2, 5).Copy .Cells(1, c + 5)
        .Cells(1, c + 5) = "第" & t_change & "次變更設計"
        .Cells(3, c).Resize(1, 5).EntireColumn.AutoFit

    End If

    Dim obj As New clsPCCES
    
    Call obj.check_item_name_repeat
    Call obj.clearOldReport(t_change)
    Call obj.getRecordingItems_export(t_change)
    Call obj.getPercentageItems_export(t_change)
    
    .Activate
    .Cells(3, c + 5).Resize(1, 5).EntireColumn.AutoFit

End With

If t_change = 0 Then Call test_getTestItems

End Sub

Sub cmdReArrange()

Dim obj As New clsPCCES

obj.ReArrangeTitle
obj.getPercentageItems

End Sub
Sub cmdFindBudget()

'Dim obj As New clsBudget
'
'obj.FindWorkbook
'If obj.IsError = True Then Exit Sub
'obj.DealBudget
'obj.clearBudget
'obj.CollectBudget
'obj.ArrangeTitle

Dim o As New clsPCCES

msg = MsgBox("變更及估驗資料確定不要了嗎?", vbYesNo + vbInformation)

If msg = vbNo Then Exit Sub

o.markTitle
o.getFileName '"D:\Users\USER\Desktop\(預算書)單期一號分線等改善工程雲林111A54_ap_bdgt.xls")
o.getAllContents
o.settingColorRules
o.getPercentageItems
o.clearMainChanges
o.clearPAY_EX

End Sub

Sub cmdShowCheckUI()

frm_Check.Show

End Sub

Sub cmdShowSingleUI()

frmData.Show

End Sub

Sub cmdShowComplexUI()

Call cmdMixComplete

MixData_Main.Show

End Sub

Sub cmdCreateProgress()

Dim o As New clsBasicData

o.addNewDiaryDays

'Dim obj2 As New clsBasicData
'
'obj2.DiartReset
'
'Dim obj As New clsInformation
'
'startDate = obj.GetStartDate
'endDate = obj.GetEndDate
'
'obj.ProgressNew

Sheets("Report").Range("C2") = "=VLOOKUP(K2,Diary!$1:$65536,2)"

End Sub

Sub cmdCreateProgressFromMain()

With ThisWorkbook.Sheets("Diary")

    Set rng = .Cells.SpecialCells(xlCellTypeLastCell)
    If rng.Row > 1 Then .Range("A2:" & rng.Address).ClearContents
    
End With

Dim o As New clsInformation

o.ProgressNew

End Sub

'Sub cmdExportToDiary()
'
'Dim obj As New clsRecord
'
'obj.cmdExportToDiary_Main
'
''obj.CollectRecDate
''obj.DealDiary
''obj.GetRecDetail
'
'End Sub
Sub cmdShowMixData()

MixData.Show vbModeless

End Sub

Sub cmdOutput()

Dim obj As New clsPrintOut

obj.BeforePrintCheck
obj.ToPDF

End Sub

Sub cmdOutput_XLS()

Dim obj As New clsPrintOut

obj.BeforePrintCheck
obj.ToXLS

End Sub

Sub cmdAddItemName()

Dim o As New clsPCCES

With Sheets("Budget")

    Set coll = o.getCollSeconedName
    
    For i = 1 To coll.Count
    
        p = p & i & "." & coll(i) & vbNewLine
    
    Next
    
    cnt = CInt(InputBox("請輸入要新增於哪個契約項次以下" & vbNewLine & p, , 1))
    
    Set rng = .Columns("B").Find(coll(cnt + 1))
    Set rng_second_name = .Columns("B").Find(coll(CInt(cnt)))
    
    .Rows(rng.Row).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    .Rows(rng.Row - 2).Copy .Rows(rng.Row - 1)
    
    item_name = InputBox("新增工項名稱=?")
    item_index = InputBox("上一編號為【" & rng.Offset(-2, -1) & "】" & vbNewLine & "新增工項編號=?", , rng_second_name.Offset(0, -1) & ".")
    item_unit = InputBox("新增工項單位=?")
    
    .Cells(rng.Row - 1, 1) = item_index
    .Cells(rng.Row - 1, 2) = item_name
    .Cells(rng.Row - 1, 3) = item_unit
    
    item_cost = InputBox("新增工項單價=?")
    
    Dim Inf_obj As New clsInformation
    
    Set coll = Inf_obj.getContractChanges

    For t_change = 0 To coll.Count - 1
    
        c = o.t_change_to_column(t_change)
    
        .Cells(rng.Row - 1, c) = 0
        .Cells(rng.Row - 1, c + 1) = item_cost
    
    Next

End With

MsgBox "新增工項完畢後請記得點選【匯出至報表】，Main的資料才會同步!", vbInformation

End Sub

Sub cmdGetPayItems(Optional pay_date As String = "")

Dim PAY_obj As New clsPay

If pay_date = "" Then pay_date = InputBox("請輸入估驗日期", , Format(Now(), "yyyy/mm/dd"))

If PAY_obj.IsPayDateLater(pay_date) = False Then
    Set coll = PAY_obj.getPayDates
    MsgBox "估驗日期需要於【" & coll(coll.Count) & "】之後!", vbCritical: End
End If

PAY_obj.pay_date = pay_date
PAY_obj.clearPAY
PAY_obj.getPayItems
PAY_obj.getOtherInf

Sheets("PAY").Visible = True
Sheets("PAY").Activate

End Sub

Sub cmdClearPayEX()

Dim myFunc As New clsMyfunction
Dim PAY_obj As New clsPay

Set coll_pay_dates = myFunc.getUniqueItems("PAY_EX", 2, "F")

i = coll_pay_dates.Count

If i = 0 Then MsgBox "查無估驗紀錄!", vbCritical: Exit Sub

pay_date = coll_pay_dates(i)

msg = MsgBox("是否要刪除最新一期【" & pay_date & "】的估驗紀錄?", vbYesNo)

If msg = vbNo Then Exit Sub

Call PAY_obj.fs_kill(i)

Set coll_rows = myFunc.getRowsByUser("PAY_EX", "F", CDate(pay_date))

For i = coll_rows.Count To 1 Step -1

    r = coll_rows(i)

    Sheets("PAY_EX").Rows(r).Delete

Next

Call PAY_obj.clearPAY

'Call cmdGetPayItems(CStr(pay_date))

MsgBox "請重新取得估驗日期!", vbCritical
'
'Sheets("Records").Activate

End Sub

Sub cmdExportToPAY()

Dim myFunc As New clsMyfunction
Dim PAY_obj As New clsPay

Set coll_rows = myFunc.getUniqueItems("PAY_EX", 2, , "估驗日期")
Set coll_pay_num = myFunc.getUniqueItems("PAY", 2, , "本次估驗")

If coll_pay_num.Count = 0 Then MsgBox "未填寫本次估驗資料，請先填寫!", vbCritical: End

PAY_obj.getPayInfo
PAY_obj.clearPAY_Report
PAY_obj.exportPayNumToReport
PAY_obj.set2ndFormula
PAY_obj.storePayItems
PAY_obj.clearPAY

'Sheets("Records").Activate

Dim Print_obj As New clsPrintOut
Dim f As String

On Error Resume Next
MkDir (getThisWorkbookPath & "\估驗Output\")
On Error GoTo 0

file_name = "第" & coll_rows.Count + 1 & "次估驗"

f = getThisWorkbookPath & "\估驗Output\" & file_name & ".xls"

'f = Application.GetSaveAsFilename(InitialFileName:="第" & coll_rows.count + 1 & "次估驗", FileFilter:="Excel Files (*.xls), *.xls")

ThisWorkbook.Sheets("PAY_Report").Visible = True

Call Print_obj.SpecificShtToXLS("PAY_Report", f) '& ".xls")
ThisWorkbook.Sheets("PAY_Report").Visible = False

ThisWorkbook.Sheets("PAY").Activate

End Sub

Sub cmdOpenPayFile()

Set fso = CreateObject("Scripting.FileSystemObject")

Dim PAY_obj As New clsPay

Dim myFunc As New clsMyfunction

Set coll_pay_dates = myFunc.getUniqueItems("PAY_EX", 2, , "估驗日期")

For i = 1 To coll_pay_dates.Count

    p = p & i & ".第" & i & "次估驗." & coll_pay_dates(i) & vbNewLine

Next

If p = "" Then
    'Sheets("Main").Range("B8") = getSavedFolder
    MsgBox "找不到已建檔的估驗資料!", vbCritical: Exit Sub
    
End If

cnt = InputBox("請輸入要打開的檔案" & vbNewLine & p, , PAY_obj.getPayCounts)

If cnt = "" Then MsgBox "未選取資料!", vbCritical: Exit Sub

If fso.FileExists(getThisWorkbookPath & "\估驗Output\" & "第" & cnt & "次估驗.xls") = True Then
    Workbooks.Open (getThisWorkbookPath & "\估驗Output\" & "第" & cnt & "次估驗.xls")
Else

    MsgBox "查無估驗資料存檔，請至儲存區看看!", vbCritical

    Shell "explorer.exe " & getThisWorkbookPath & "\" & "PAY\", vbNormalFocus
    
End If

End Sub

Sub cmdGetProgressByInterpolation()

Dim o As New clsBasicData

Call o.getProgByInter

End Sub

Sub cmdStopProgress()

Dim o As New clsBasicData

Call o.addStopDays

Sheets("Diary").Activate

End Sub

Sub cmdEnlargeWorkDays()

On Error GoTo ERRORHANDLE

Dim Inf_obj As New clsInformation
Dim myFunc As New clsMyfunction

enlargeDate = CDate(InputBox("請輸入展延開始日期", , Format(Now(), "yyyy/mm/dd")))
enlargeDays = CInt(InputBox("請輸入展延天數", , 1))

Sheets("Main").Range("B6") = Inf_obj.workDay + enlargeDays
Sheets("Main").Range("C6") = CDate(enlargeDate)

With Sheets("Diary")

    .Activate

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row

    For i = 1 To enlargeDays
        
        end_date = Inf_obj.GetEndDate
        diary_date = end_date + i
        
        Call myFunc.AppendData("Diary", Array(lr + i - 1, diary_date, "晴"))
        
        '----set formula---
        
        .Cells(lr + i, 1).Resize(1, 10).Borders.LineStyle = 1
        .Cells(lr + i, 1).Resize(1, 4).HorizontalAlignment = xlCenter
        .Cells(lr + i, 2).NumberFormatLocal = "yyyy/mm/dd(aaa)"
        .Cells(lr + i, 5).Resize(1, 2).WrapText = True
        .Cells(lr + i, 4).NumberFormatLocal = "0.00%"
        
        If i = enlargeDays Then
        
            .Cells(lr + i, 4) = 1
            .Cells(lr, 4) = ""
        
        End If
        
    Next

End With

'msg = MsgBox("是否要自展延的最後一天前預定進度設定?", vbYesNo)
'
'If msg = vbYes Then
'
'    Call cmdGetProgressByInterpolation
'
'End If

Exit Sub

ERRORHANDLE:

MsgBox "動作已取消!", vbCritical

End Sub

Sub checkTestCompleted(Optional ByVal IsUpload As Boolean = False) '20230225 add

Dim REC_obj As New clsRecord
Dim Test_obj As New clsReportTest

If IsUpload = True Then

Dim GAS_obj As New clsFetchURL_TEST

End If

With Sheets("Test")

    lr = .Cells(.Rows.Count, "C").End(xlUp).Row
    
    For r = 2 To lr
    
        TestName = .Cells(r, "A")
        ItemName = .Cells(r, "C")
        testPeriod = .Cells(r, "D")
        
        Call REC_obj.getNumAndSumByItemName(TestName, CDate(Now()), rec_num, rec_sum)
        
        ItemName = .Cells(r, "C")
        
        Call REC_obj.getNumAndSumByItemName(ItemName, CDate(Now()), item_num, item_sum)
        
        calcTest = rec_sum
        doTest = CInt(Test_obj.getTestNeedNum(item_sum, testPeriod))
        
        If doTest > calcTest Then
        
            prompt = prompt & TestName & "尚欠缺" & doTest - calcTest & "組" & vbNewLine & vbNewLine
        
        End If
        
        If IsUpload = True Then
        
        myURL = GAS_obj.CreateURL(TestName, doTest - calcTest)
        GAS_obj.ExecHTTP (myURL)
        
        End If
    
    Next

    If prompt <> "" Then MsgBox prompt

End With

End Sub

Sub cmdGetReportSum()

Dim obj As New clsRecord

obj.getReportSum

Dim Print_obj As New clsPrintOut

ThisWorkbook.Sheets("Report_Sum").Visible = True

Print_obj.SpecificShtToXLS ("Report_Sum")

ThisWorkbook.Sheets("Report_Sum").Visible = False

End Sub

Sub deleteRecIndex()

rec_Index = InputBox("請輸入預計要刪除的流水號:")

Dim myFunc As New clsMyfunction

Set coll_rows = myFunc.getRowsByUser2("Records", rec_Index, 2, "流水號")

If coll_rows.Count = 0 Then MsgBox "查無此流水號!", vbCritical: End

Set coll_rows = myFunc.ReverseColl(coll_rows)

For Each r In coll_rows
    
    Sheets("Records").Rows(r).Delete

Next

End Sub

Sub cmdGetCheckTable() '檢驗停留點申請單

Dim Print_obj As New clsPrintOut
Dim Inf_obj As New clsInformation
Dim myFunc As New clsMyfunction

Set checkdaylist = myFunc.getUniqueItems("Check", 3, , "時間") ' getTimeList

For Each checkday In checkdaylist

    cnt = cnt + 1
    
    p = p & cnt & "." & checkday & vbNewLine

Next

j = InputBox("請問要列印哪一天?" & vbNewLine & p, , 1)

Show_CheckDate = checkdaylist(CInt(j))

msg = MsgBox("是否要將停留點申請單全部更新?", vbInformation + vbYesNo)

If msg = vbYes Then

    With Sheets("Check")
    
    lr = .Cells(2, 1).End(xlDown).Row
    
    For Each checkday In checkdaylist
    
    myRow = 15
    
    i = i + 1
    
    With Sheets("CheckList")
     
        .Range("W4") = i
        .Range("W6") = CDate(checkday) - 1
        .Cells(15, 1).Resize(10, 26).ClearContents
    
    End With
    
        For r = 2 To lr
            
            If .Cells(r, 4) = CDate(checkday) And .Cells(r, 5) = "檢驗停留點" Then
            
                checkitem = .Cells(r, 1)
                tmp = split(.Cells(r, 6), ",")
                checkch = tmp(0)
                checkloc = tmp(1)
            
                With Sheets("CheckList")
                    
                    .Range("E8") = Inf_obj.conName
                    .Range("E10") = Inf_obj.contractor
                    .Range("A" & myRow) = checkch
                    .Range("G" & myRow) = checkday
                    .Range("M" & myRow) = checkloc
                    .Range("R" & myRow) = checkitem
                
                    myRow = myRow + 1
                
                End With
            
            End If
            
        Next
    
        If myRow = 15 Then
        
            i = i - 1
            
        Else
        
            Sheets("CheckList").Visible = True
            
            Call Print_obj.SpecificShtToXLS("CheckList", getThisWorkbookPath & "\抽查表Output\EN-" & i & ".xls", "EN-" & i)
            
            Sheets("CheckList").Visible = False
            
        End If
    
    Next
    
    End With

End If

If myFunc.IsFileExists(getThisWorkbookPath & "\抽查表Output\EN-" & j & ".xls") = True Then

Workbooks.Open (getThisWorkbookPath & "\抽查表Output\EN-" & j & ".xls")

Else

MsgBox "找不到該日報表，請全部更新後再試一次!", vbCritical

End If

'Call SendEmail("apple84026113@gmail.com", getThisWorkbookPath & "\抽查表Output\EN-" & j & ".xls")

'MsgBox "已將信件發送給主辦提請抽查!", vbInformation

End Sub

Sub cmdMergeChecks()

Dim coll_rows As New Collection

For Each rng In Selection

    If rng.Row > 2 Then
    
        On Error Resume Next
        coll_rows.Add rng.Row, CStr(rng.Row)
        On Error GoTo 0
        
    End If

Next

Dim myFunc As New clsMyfunction

Dim coll As New Collection

With Sheets("Check")

    For Each r In coll_rows
    
        check_index = .Cells(r, 2)
        check_num = .Cells(r, 3)
        check_date = .Cells(r, 4)
        
        check_inf = .Cells(r, 7)
        photo_inf = .Cells(r, 9)
        
        If check_inf <> "" Then
        
            check_path = getThisWorkbookPath & "\抽查表Output\" & check_index & "-" & check_num & ".xls"
            coll.Add check_path
        
        End If
        
        If photo_inf <> "" Then
        
            photo_path = getThisWorkbookPath & "\查驗照片Output\" & check_index & "-" & check_num & ".xls"
            
            If myFunc.IsFileExists(photo_path) = False Then
            
                Call PastePhoto(r, True)
                coll.Add photo_path
            Else
            
                coll.Add photo_path
            
            End If
        
        End If
        
    Next

End With

Dim Print_obj As New clsPrintOut

Call Print_obj.combineFiles(coll)

End Sub

Sub cmdPastePhotos()

Dim myFunc As New clsMyfunction
Dim FilePath As String
Dim IsXLS As Boolean

mode_msg = MsgBox("是否列印PDF?", vbInformation + vbYesNo)

If mode_msg = vbYes Then
    IsXLS = False
Else
    IsXLS = True
End If

With Sheets("Check")

For Each rng In Selection

    r = rng.Row

    If r > 2 Then
    
        check_index = .Cells(r, 2)
        check_num = .Cells(r, 3)
    
        FilePath = getThisWorkbookPath & "\查驗照片Output\" & check_index & "-" & check_num & ".xls"
    
        If myFunc.IsFileExists(FilePath) = True Then
        
            msg = MsgBox("查驗照片Output中已存在【" & check_index & "-" & check_num & "】，是否要取代?", vbYesNo + vbInformation)
            
            If msg = vbYes Then Call PastePhoto(r, IsXLS)
            
        Else
        
            Call PastePhoto(r, IsXLS)
        
        End If
    
    End If

Next

End With

End Sub

Sub PastePhoto(ByVal r As Integer, ByVal IsXLS As Boolean)

Dim o As New clsReportPhoto
Dim Inf_obj As New clsInformation
Dim myFunc As clsMyfunction

Sheets("ReportPhoto").Range("A1") = Inf_obj.conName

o.IsXLS = IsXLS

''If IsAsk = True Then
'
'    'msg = MsgBox("是否列印PDF?", vbYesNo)
'
'    If msg = vbYes Then
'        o.IsXLS = False
'    Else
'        o.IsXLS = True
'    End If
'
'Else
'
'o.IsXLS = True

'End If

If Sheets("Check").Range("E1") = "Y" Then
    o.IsShowText = True
Else
    o.IsShowText = False
End If

With Sheets("Check")

    'lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    'For r = 3 To lr
    
        check_name = .Cells(r, 1)
        check_eng = .Cells(r, 2)
        check_num = .Cells(r, 3)
        
        check_photo_inf = .Cells(r, "I")
        
        If check_photo_inf <> "" Then
        
            Call o.GetReportByItem(r)
        
        End If
    
    'Next
    
.Activate

End With

End Sub


Sub cmdEditCheck()

On Error GoTo ERRORHANDLE
r = Selection.Row

If r < 3 Then GoTo ERRORHANDLE

On Error GoTo 0

With Sheets("Check")

    .Rows(r).Select
    
    check_name = .Cells(r, 1)
    check_eng = .Cells(r, 2)
    check_num = .Cells(r, 3)
    check_date = .Cells(r, 4)
    check_style = .Cells(r, 5)
    check_loc = .Cells(r, 6)
    photo_lst = splitPhotoList(.Cells(r, "I"))
    
    If check_name = "" Then GoTo ERRORHANDLE
    
    msg = MsgBox("請問是否要編輯[" & check_eng & "]-" & check_num & "?", vbInformation + vbYesNo)
    
    If msg = vbNo Then
        Exit Sub
    Else

        With frm_Check
        
            .Label_Row = CInt(r)
            .cboCheckItem.Value = check_name & "[" & check_eng & "]"
            .cboCheckItem.Enabled = False
            .txtCheckDate = check_date
            .txtCheckDate.Enabled = False
            .cboCheckStyle.Value = check_style
            tmp = split(check_loc, ",")
            .txtCheckCanal = tmp(0)
            .txtCheckLocDetail = tmp(1)
            .txtCheckLoc = check_loc
            
            If IsEmpty(photo_lst) = False Then
      
                For i = LBound(photo_lst) To UBound(photo_lst)
                
                    .lstCheckTable.AddItem ""
                    .lstCheckTable.List(i, 0) = photo_lst(i, 0)
                    .lstCheckTable.List(i, 1) = photo_lst(i, 1)
                
                Next
                
            End If
            
            '.lstCheckTable = splitPhotoList(photo_lst)
            
            .Show
            
        End With
        
    End If

End With

Exit Sub

ERRORHANDLE:

MsgBox "請框選要編修的儲存格!", vbInformation

End Sub

Sub cmdDeleteCheckBySelect()

On Error GoTo ERRORHANDLE

Dim coll As New Collection

For Each rng In Selection

    r = rng.Row
    
    On Error Resume Next
    
    coll.Add r, CStr(r)
    
    On Error GoTo 0

Next

Dim myFunc As New clsMyfunction

Call myFunc.BubbleSort_coll(coll)

Set coll_rows = myFunc.ReverseColl(coll)

With Sheets("Check")

For Each r In coll_rows
    
    MyCode = .Cells(r, 2)
    myNum = .Cells(r, 3)
    
    myFilePath = getThisWorkbookPath & "\查驗照片Output\" & MyCode & "-" & myNum & ".xls"
    
    If myFunc.IsFileExists(myFilePath) = True Then Kill myFilePath
    
    myFilePath = getThisWorkbookPath & "\抽查表Output\" & MyCode & "-" & myNum & ".xls"
    
    If myFunc.IsFileExists(myFilePath) = True Then Kill myFilePath
    
    .Rows(r).Delete
    
Next

End With

Exit Sub
ERRORHANDLE:
MsgBox "請框選要刪除的欄位!", vbCritical

End Sub


'=============function===============

Function getRemainedItems(ByVal rec_date As Date)

Dim myFunc As New clsMyfunction
Dim PCCES_obj As New clsPCCES
Dim Inf_obj As New clsInformation
Dim REC_obj As New clsRecord
Dim coll_Need As New Collection

Set coll_item_names = PCCES_obj.getRecordingItemsByRecDate(rec_date)
t_change = Inf_obj.getContractChangesByDate(rec_date)

For i = 1 To coll_item_names.Count

    item_name = coll_item_names(i)

    Set coll_rows = myFunc.getRowsByUser("Budget", "B", item_name)
    
    contract_num = Sheets("Budget").Cells(coll_rows(1), PCCES_obj.t_change_to_column(t_change))
    
    Call REC_obj.getNumAndSumByItemName(item_name, rec_date, rec_num, rec_sum)
    
    If contract_num - rec_sum <> 0 Then coll_Need.Add item_name

Next

Set getRemainedItems = coll_Need

End Function

Function getSumByItemNameAndCanal(ByVal item_name As String, ByVal canal_name As String)

Dim f As New clsMyfunction

'Set coll_rows = f.getRowsByUser("TMP", "B", item_name)

Dim coll_rows As New Collection

With Sheets("TMP")

    For r = 2 To .Cells(.Rows.Count, 1).End(xlUp).Row
    
        If .Cells(r, "B") = item_name Then coll_rows.Add r
    
    Next

End With


For Each r In coll_rows

    canalName = Sheets("TMP").Cells(r, 1)
    rec_sum = Sheets("TMP").Cells(r, 3)
    
    If canalName = canal_name Then getSumByItemNameAndCanal = rec_sum

Next

End Function

Function getThisWorkbookPath()

Set fso = CreateObject("Scripting.FileSystemObject")

myPath = ThisWorkbook.Sheets("Main").Range("B8")
exePath = ThisWorkbook.Path

If fso.FolderExists(exePath & "\監造日報表Output\") = True Then

ThisWorkbook.Sheets("Main").Range("B8") = exePath

ElseIf fso.FolderExists(myPath & "\監造日報表Output\") = True Then

ThisWorkbook.Sheets("Main").Range("B8") = myPath

Else

ThisWorkbook.Sheets("Main").Range("B8") = getSavedFolder

End If

getThisWorkbookPath = ThisWorkbook.Sheets("Main").Range("B8")

End Function


