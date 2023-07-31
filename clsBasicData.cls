VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBasicData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private collChannel As New Collection
Private collTest As New Collection
Private collitem As New Collection
Private collItemRule As New Collection

Private shtMain As Object
Private shtRec As Object
Private shtBudget As Object
Private shtReport As Object
Private shtDiary As Object

Private c_Main As Integer
Private ConRow As Integer
Private TestRow As Integer
Private ChannelRow As Integer
Private ItemRow As Integer
Private PgsRow As Integer

Private ReportDay As Date
Private t_change As Integer

Private Sub Class_Initialize()

ThisWorkbook.Activate

Set shtMain = Sheets("Main")
Set shtRec = Sheets("Records")
Set shtBudget = Sheets("Budget")
Set shtDiary = Sheets("Diary")
Set shtReport = Sheets("Report")

'c_Main = 6

Dim obj As New clsInformation

If IsError(shtReport.Range("C2")) = False Then

ReportDay = CDate(shtReport.Range("C2"))

t_change = obj.getContractChangesByDate(ReportDay)

c_Main = 6 + t_change * 5

Else

c_Main = 6

End If

'Set collContractDay = obj.GetcollContractDay
'
'ReportDay = shtReport.Range("C2")
'
'If collContractDay.count > 1 Then
'
'    For i = collContractDay.count To 2 Step -1
'
'        If ReportDay >= collContractDay.Item(i) Then
'            c_Main = 6 + (i - 1) * 5
'            t_change = t_change + 1
'            Exit For
'        End If
'    Next
'
'End If

End Sub

Sub ReadData() '讀取工程項目、試驗項目、施作位置 for frmData

rec_date = frmData.txtDay

Dim PCCES_obj As New clsPCCES

Call GetMainRow

With shtMain
    
    Set collitem = PCCES_obj.getRecordingItemsByRecDate(CDate(rec_date))
    
    '----about t_change------
    
'    lr = .Cells(ItemRow, c_Main).End(xlDown).Row
'
'    For r = ItemRow + 1 To lr
'
'        collItem.Add .Cells(r, c_Main).Value
'        collItemRule.Add .Cells(r, c_Main).Font.ColorIndex
'
'    Next
'
    '==========================
    
    For r = TestRow + 1 To ChannelRow - 1
    
        If .Cells(r, 1) <> "" Then collTest.Add .Cells(r, 1).Value
    
    Next
    
    For r = ChannelRow + 1 To PgsRow - 1
    
        If .Cells(r, 1) <> "" Then collChannel.Add .Cells(r, 1).Value
    
    Next

End With

End Sub

Sub DiaryReset() '重整Diary工作表

'Sub DealDiary() '初始化日誌

With shtDiary

    lr = .Cells(Rows.Count, 1).End(xlUp).Row

    .Activate

    For r = 2 To lr
    
        'If r = 14 Then Stop
    
        For c = 5 To 6
    
            If .Cells(r, c).Interior.ColorIndex = -4142 Then .Cells(r, c).ClearContents '網底變色即鎖住
    
        Next
    
    Next

End With

'End Sub

Exit Sub

Dim obj As New clsInformation

startDate = obj.GetStartDate
endDate = obj.GetEndDate

Count = endDate - startDate + 1

With shtDiary

    If .Cells(2, 1) = startDate Then Exit Sub

    MyRowH = 25
    
    lr = .Cells(Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To Count + 1
        
        'If r = 110 Then
        
        'ElseIf r > 110 Then
        
        'End If
        myAddDay = .Cells(r, "J")
        .Cells(r, 1) = r - 1
        .Cells(r, 2) = startDate + r - 2 + myAddDay
        
        .Rows(r).RowHeight = MyRowH
        
        If .Cells(r, 5) <> "" Then .Rows(r).AutoFit
        
        If .Rows(r).RowHeight < MyRowH Then .Rows(r).RowHeight = MyRowH
    
    Next

    Do Until lr <= Count + 1
    
        .Rows(r).Delete
    
        lr = lr - 1
    
    Loop
    
End With


End Sub

Sub Init() '於frmData表單新增施作位置、工程項目、試體項目

With frmData

    For i = 1 To collChannel.Count
        
        .cboChannel.AddItem collChannel.Item(i)
    
    Next
    
    For i = 1 To collitem.Count '設定不顯示的百分比項目
    
        'If collItemRule.Item(i) <> 5 Then
        .cboItem.AddItem collitem.Item(i)
    
    Next
    
    For i = 1 To collTest.Count
    
        .cboItem.AddItem collTest.Item(i)
    
    Next

    '.txtDay = Format(Now(), "yyyy/mm/dd")

End With

End Sub

Sub Init_Mix() '於MixData表單新增工程項目、試體項目

Dim PCCES_obj As New clsPCCES
Dim Inf_obj As New clsInformation

Set collitem = PCCES_obj.getRecordingItemsAll

With MixData

    For i = 1 To collitem.Count '設定不顯示的百分比項目
    
        'If collItemRule.Item(i) <> 5 Then
        .cboItem.AddItem collitem.Item(i)
    
    Next
    
'    For i = 1 To collTest.count
'
'        .cboItem.AddItem collTest.Item(i)
'
'    Next
    
End With

End Sub

Sub RetrunUnit() '查詢TargetItem對應剩餘數、單位、預儲存未扣除之數量

Call GetMainRow

TargetItem = frmData.cboItem

With shtMain

    For r = ItemRow + 1 To .Cells(Rows.Count, c_Main).End(xlUp).Row
    
        If .Cells(r, c_Main) = TargetItem Then
        
            frmData.lblUnit = .Cells(r, c_Main + 1)
            frmData.lblLast = .Cells(r, c_Main + 2)
            frmData.lblStore = .Cells(r, c_Main + 2)

            Exit Sub
        
        End If
    
    Next
    
    For r = TestRow + 1 To ChannelRow - 1
    
        If TargetItem = .Cells(r, 1) Then
        
            frmData.lblUnit = .Cells(r, 3) '將試驗單價改為試驗單位
            frmData.lblLast = .Cells(r, 2 + t_change * 3)
            frmData.lblStore = .Cells(r, 2 + t_change * 3)
            
        End If
        
    Next

End With

End Sub

Sub RetrunUnit_Mix() '查詢TargetItem對應剩餘數、單位

Call GetMainRow

TargetItem = MixData.cboItem

With shtMain

    For r = ItemRow + 1 To .Cells(Rows.Count, c_Main).End(xlUp).Row
    
        If .Cells(r, c_Main) = TargetItem Then
        
            MixData.lblUnit = .Cells(r, c_Main + 1)

            Exit Sub
        
        End If
    
    Next
    
    For r = TestRow + 1 To ChannelRow - 1
    
        If TargetItem = .Cells(r, 1) Then
        
            frmData.lblUnit = .Cells(r, 3)
            
        End If
        
    Next

End With

End Sub

Sub UsedAmount() '查詢已經使用的數量

Dim UntilDay As Date

With frmData

TargetItem = .cboItem
UntilDay = .txtDay
IsUntilDay = .chkUntilDay

End With

With shtRec

    For r = 3 To .Cells(3, 1).End(xlDown).Row
    
        RecDay = .Cells(r, 2)
        
        If RecDay > UntilDay And IsUntilDay = True Then Exit For
        
        RecItem = .Cells(r, 5)
        RecAmount = .Cells(r, 6)
        
        If RecItem = TargetItem Then sumAmount = sumAmount - RecAmount
    
    Next

End With

With frmData '這裡有bug，210kg/cm2混凝土所得結果=1.421*E^-14

.lblStore = Round(.lblStore, 5) + Round(sumAmount, 5)

.lblLast = .lblStore

End With

End Sub

Sub getProgByInter() '20230624

Call checkProgSetting
Set collProg = getProgColl

With Sheets("Diary")

    lr = .Cells(.Rows.Count, 4).End(xlUp).Row

    For r = 2 To lr
    
        myProg = .Cells(r, 4)
        
        If myProg = "" Then
        
            For i = 1 To collProg.Count
                
                tmp = split(collProg(i), ":")
                
                If r <= CInt(tmp(0)) Then
                
                    r1 = split(collProg(i - 1), ":")(0)
                    p1 = split(collProg(i - 1), ":")(1)
                    
                    r2 = split(collProg(i), ":")(0)
                    p2 = split(collProg(i), ":")(1)
                    
                    newProg = Round(((r2 - r) * p1 + (r - r1) * p2) / (r2 - r1), 4)
                    
                    Exit For
                
                End If
            
            Next
            
            .Cells(r, 4) = newProg
        
        End If
        
    Next

End With

End Sub

Sub addNewDiaryDays()

cnt = InputBox("請輸入工期")

Sheets("Main").Range("B5") = cnt

Dim myFunc As New clsMyfunction
Dim Inf_obj As New clsInformation

'Inf_obj.getBasicInformation
startDate = Inf_obj.startDate

Call clearDiary

With Sheets("Diary")

    .Activate

    For i = 1 To cnt
    
        diary_date = startDate + i - 1
        
        Call myFunc.AppendData("Diary", Array(i, diary_date, "晴"))
        
        '----set formula---
        
        .Cells(i + 1, 1).Resize(1, 10).Borders.LineStyle = 1
        .Cells(i + 1, 1).Resize(1, 4).HorizontalAlignment = xlCenter
        .Cells(i + 1, 2).NumberFormatLocal = "yyyy/mm/dd(aaa)"
        .Cells(i + 1, 5).Resize(1, 2).WrapText = True
        .Cells(i + 1, 4).NumberFormatLocal = "0.00%"
        
    Next

End With

End Sub

Sub clearDiary()

With ThisWorkbook.Sheets("Diary")

    Set rng = .Cells.SpecialCells(xlCellTypeLastCell)
    If rng.Row > 1 Then .Range("A2:" & rng.Address).Clear
    
End With

End Sub

Sub addStopDays()

Dim myFunc As New clsMyfunction

Set coll_rows = myFunc.getUniqueItems("Diary", 2, , "預定進度")

If coll_rows.Count = 0 Then MsgBox "請先輸入預定進度資料!", vbCritical: End

stop_date_begin = InputBox("請輸入起始停工日期", , "2023/8/1")
stop_date_end = InputBox("請輸入結束停工日期", , "2023/8/6")

cnt = CDate(stop_date_end) - CDate(stop_date_begin) + 1

With Sheets("Diary")

    Set rng = .Columns("B").Find(Format(CDate(stop_date_end), "yyyy/mm/dd(aaa)"))
    If rng Is Nothing Then MsgBox "找不到" & stop_date_end & "!", vbCritical: End
    
    Set rng = .Columns("B").Find(Format(CDate(stop_date_begin), "yyyy/mm/dd(aaa)"))
    If rng Is Nothing Then MsgBox "找不到" & stop_date_begin & "!", vbCritical: End
    
    

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
        
    For i = 1 To cnt
    
        .Cells(lr + i, "A") = .Cells(lr + i - 1, "A") + 1
        .Cells(lr + i, "B") = .Cells(lr + i - 1, "B") + 1
        .Cells(lr + i, "C") = "晴"
        
        .Cells(lr + i, 1).Resize(1, 10).Borders.LineStyle = 1
        .Cells(lr + i, 1).Resize(1, 4).HorizontalAlignment = xlCenter
        .Cells(lr + i, 2).NumberFormatLocal = "yyyy/mm/dd(aaa)"
        .Cells(lr + i, 5).Resize(1, 2).WrapText = True
        .Cells(lr + i, 4).NumberFormatLocal = "0.00%"
        
    Next
    
    Set coll_rows = myFunc.getRowsByUser2("Diary", Format(stop_date_begin, "yyyy/mm/dd(aaa)"), 1, "報表日期")

    r_begin = coll_rows(1)
    first_pgs = .Cells(r_begin - 1, "D")
    
    .Range("D" & r_begin & ":D" & lr).Copy .Range("D" & r_begin + cnt)

    For r = r_begin To r_begin + cnt - 1
    
        .Cells(r, "D") = .Cells(r - 1, "D")
        .Cells(r, "D").Interior.ColorIndex = 22
    
    Next
    
End With

End Sub

'-----------FUNCTION-----------------

Sub checkProgSetting()

Dim Inf_obj As New clsInformation

'Inf_obj.getBasicInformation

fixStartDate = Inf_obj.startDate

With Sheets("Diary")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    progStartDate = .Cells(2, 2)
    progStartProg = .Cells(2, 4)
    progEndDate = .Cells(lr, 2)
    progEndProg = .Cells(lr, 4)
    
    If progStartDate <> fixStartDate Then
    
        MsgBox ("開工日「" & progStartDate & "」，與標案設定開工日「" & fixStartDate & "」不一樣!"), vbCritical
        End
        
    End If
    
    If progStartProg = "" Then
    
        .Cells(2, 4) = 0
        MsgBox "系統自動於開工日補上0%", vbInformation
        
    End If
    
    If progEndProg <> 1 Then
    
        .Cells(lr, 4) = 1
        MsgBox "系統自動於竣工日補上100%", vbInformation
    
    End If
    
End With

End Sub

Private Function getProgColl()

Dim coll As New Collection

With Sheets("Diary")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row

    '------main-----------
    
    For r = 2 To lr
    
        mydate = .Cells(r, 2)
        myProg = .Cells(r, 4)
        
        If myProg <> "" Then
        
            coll.Add r & ":" & myProg
        
        End If
    
    Next
    
    For i = 1 To coll.Count - 1
    
        sr = split(coll(i), ":")(0)
        b_pgs = split(coll(i), ":")(1)
        er = split(coll(i + 1), ":")(0)
        e_pgs = split(coll(i + 1), ":")(1)
    
        If b_pgs > e_pgs Then
        
            MsgBox "【" & .Cells(sr, 2) & "】進度不能比【" & .Cells(er, 2) & "】快!", vbCritical: End
    
        End If
    
    Next
    
End With

Set getProgColl = coll

If coll.Count = 2 Then MsgBox ("建議在預定進度的欄位「D」填寫進度，內差成果才會比較準確!"), vbCritical

End Function


Private Sub GetMainRow()

Call ReturnMainRow(arr)

ConRow = arr(0)
TestRow = arr(1)
ChannelRow = arr(2)
ItemRow = arr(3)
PgsRow = arr(4)

End Sub



