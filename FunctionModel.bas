Attribute VB_Name = "FunctionModel"
Sub cmdGetReportIDByDate(Optional myNewDate As Date) '20221125依日期選擇頁數

With Sheets("Report")

    mydate = .Range("C2")
    myID = .Range("K2")
    
    If myNewDate = 0 Then myNewDate = InputBox("請輸入日期，格式如" & vbNewLine & mydate, , mydate)
    On Error GoTo DATEFORMATERRORHANDLE
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
    
        conNum = .Cells(r, "F")
        sumNum = .Cells(r, "I")
        
        If conNum <> sumNum Then
        
            ItemName = .Cells(r, "B")
            numDiff = Round(sumNum - conNum, 4)
            
            If Abs(numDiff) < allowence Then
            
                Call dealOverNum(ItemName, numDiff)
            
                prompt = prompt & vbNewLine & ItemName & ":" & numDiff
        
            End If
        
        End If
    
    Next
    
    MsgBox prompt, vbInformation

End With

End Sub

Sub dealOverNum(ByVal ItemName As String, ByVal numDiff As Double) '20221122處理剩餘零星數量

With Sheets("Records")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row
    
    For r = lr To 3 Step -1
    
        recName = .Cells(r, "E")
        
        If recName = ItemName Then
        
            originNum = .Cells(r, "F")
            
            adjustNum = originNum - numDiff
            
            If adjustNum > 0 Then
            
                Debug.Print ItemName & ",原數量=" & originNum & ">>校正=" & adjustNum
            
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

lr = .Cells(.Rows.count, "D").End(xlUp).Row

For r = 3 To lr

    If .Cells(r, 1) <> "" Then
    
        mix_name = .Cells(r, 1)
        
        Call o.ChangeMixToRecord(mix_name)
        
    End If

Next

End With

End Sub

Sub test_getMixSumUnit()

mylen = InputBox("單元總長=?")
myName = InputBox("單元名稱")

Dim coll As New Collection

With Sheets("Mix_Sum")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row
    
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
    
        lr = .Cells(.Rows.count, 1).End(xlUp).Row + 1
        
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

'Sub cmdgetTmpData()
'
'Dim obj As New clsRecord
'
''msg = MsgBox("是否要留白?", vbYesNo)
'
''If msg = vbYes Then
'obj.getTmpData (True)
''Else
''obj.getTmpData (False)
''End If
'
'obj.Tmp2TmpTotal
'
'End Sub

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

Sub cmdPrintCheck()

Dim obj As New clsCheck

'obj.CheckList
obj.PrintCheckTable

End Sub

Sub cmdExportToCheck()

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

End Sub

Sub cmdReArrange()

Dim obj As New clsPCCES

obj.ReArrangeTitle

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

Sub cmdShowSingleUI()

frmData.Show

End Sub

Sub cmdShowComplexUI()

MixData_Main.Show

End Sub

Sub cmdCreateProgress()

Dim obj2 As New clsBasicData

obj2.DiartReset

Dim obj As New clsInformation

startDate = obj.GetStartDate
endDate = obj.GetEndDate

obj.ProgressNew

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
    
    For i = 1 To coll.count
    
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

    For t_change = 0 To coll.count - 1
    
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
    MsgBox "估驗日期需要於【" & coll(coll.count) & "】之後!", vbCritical: End
End If

PAY_obj.pay_date = pay_date
PAY_obj.clearPAY
PAY_obj.getPayItems
PAY_obj.getOtherInf

Sheets("PAY").Activate

End Sub

Sub cmdClearPayEX()

Dim myFunc As New clsMyfunction
Dim PAY_obj As New clsPay

Set coll_pay_dates = myFunc.getUniqueItems("PAY_EX", 2, "F")

i = coll_pay_dates.count

If i = 0 Then MsgBox "查無估驗紀錄!", vbCritical: Exit Sub

pay_date = coll_pay_dates(i)

msg = MsgBox("是否要刪除最新一期【" & pay_date & "】的估驗紀錄?", vbYesNo)

If msg = vbNo Then Exit Sub

Call PAY_obj.fs_kill(i)

Set coll_Rows = myFunc.getRowsByUser("PAY_EX", "F", CDate(pay_date))

For i = coll_Rows.count To 1 Step -1

    r = coll_Rows(i)

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

Set coll_Rows = myFunc.getUniqueItems("PAY_EX", 2, , "估驗日期")
Set coll_pay_num = myFunc.getUniqueItems("PAY", 2, , "本次估驗")

If coll_pay_num.count = 0 Then MsgBox "未填寫本次估驗資料，請先填寫!", vbCritical: End

PAY_obj.getPayInfo
PAY_obj.clearPAY_Report
PAY_obj.exportPayNumToReport
PAY_obj.set2ndFormula
PAY_obj.storePayItems
PAY_obj.clearPAY

Sheets("Records").Activate

Dim Print_obj As New clsPrintOut
Dim f As String

On Error Resume Next
MkDir (ThisWorkbook.Path & "\PAY\")
On Error GoTo 0

file_name = "第" & coll_Rows.count + 1 & "次估驗"

f = ThisWorkbook.Path & "\PAY\" & file_name & ".xls"

'f = Application.GetSaveAsFilename(InitialFileName:="第" & coll_rows.count + 1 & "次估驗", FileFilter:="Excel Files (*.xls), *.xls")

ThisWorkbook.Sheets("PAY_Report").Visible = True

Call Print_obj.SpecificShtToXLS("PAY_Report", f) '& ".xls")
ThisWorkbook.Sheets("PAY_Report").Visible = False

End Sub

Sub cmdOpenPayFile()

Set fso = CreateObject("Scripting.FileSystemObject")

Dim PAY_obj As New clsPay

Dim myFunc As New clsMyfunction

Set coll_pay_dates = myFunc.getUniqueItems("PAY_EX", 2, , "估驗日期")

For i = 1 To coll_pay_dates.count

    p = p & i & ".第" & i & "次估驗." & coll_pay_dates(i)

Next

If p = "" Then MsgBox "找不到已建檔的估驗資料!", vbCritical: Exit Sub

cnt = InputBox("請輸入要打開的檔案" & vbNewLine & p, , PAY_obj.getPayCounts)

If fso.FileExists(ThisWorkbook.Path & "\PAY\" & "第" & cnt & "次估驗.xls") = True Then

    Workbooks.Open (ThisWorkbook.Path & "\PAY\" & "第" & cnt & "次估驗.xls")
Else

Shell "explorer.exe " & wbpath & "\" & "PAY\", vbNormalFocus
    
End If

End Sub

'=============function===============

Function getRemainedItems(ByVal rec_date As Date)

Dim myFunc As New clsMyfunction
Dim PCCES_obj As New clsPCCES
Dim Inf_obj As New clsInformation
Dim REC_obj As New clsRecord
Dim coll_Need As New Collection

'rec_date = CDate("2023/7/17")

Set coll_item_names = PCCES_obj.getRecordingItemsByRecDate(rec_date)
t_change = Inf_obj.getContractChangesByDate(rec_date)

For i = 1 To coll_item_names.count

    item_name = coll_item_names(i)

    Set coll_Rows = myFunc.getRowsByUser("Budget", "B", item_name)
    
    contract_num = Sheets("Budget").Cells(coll_Rows(1), PCCES_obj.t_change_to_column(t_change))
    
    Call REC_obj.getNumAndSumByItemName(item_name, rec_date, rec_num, rec_sum)
    
    If contract_num - rec_sum <> 0 Then coll_Need.Add item_name

Next

Set getRemainedItems = coll_Need

End Function

Function getTestNeedNum(ByVal num As Double, ByVal s As String)

tmp = Split(s, ",")

For Each it In tmp

    If IsNumeric(it) Then

        If num >= CDbl(it) Then cnt = cnt + 1
    
    Else
    
        If cnt > 0 Then
    
            before_num = CDbl(tmp(j - 1))
            each_num = CDbl(mid(it, 1, Len(it) - 1))
         
            If (num - before_num) <> 0 Then
        
            cnt = cnt + Int((num - before_num) / each_num) + 1
    
            End If
    
        End If
    
    End If
    
    j = j + 1
    
Next

getTestNeedNum = cnt

End Function

