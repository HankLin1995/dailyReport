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
Private collItem As New Collection
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

Private t_change As Integer

Private Sub Class_Initialize()

ThisWorkbook.Activate

Set shtMain = Sheets("Main")
Set shtRec = Sheets("Records")
Set shtBudget = Sheets("Budget")
Set shtDiary = Sheets("Diary")
Set shtReport = Sheets("Report")

c_Main = 6

Dim obj As New clsInformation

Set collContractDay = obj.GetcollContractDay

ReportDay = shtReport.Range("C2")

If collContractDay.count > 1 Then

    For i = collContractDay.count To 2 Step -1
    
        If ReportDay >= collContractDay.Item(i) Then
            c_Main = 6 + (i - 1) * 5
            t_change = t_change + 1
            Exit For
        End If
    Next
    
End If

End Sub

Sub ReadData() '讀取工程項目、試驗項目、施作位置

Call GetMainRow

With shtMain
    
    lr = .Cells(ItemRow, c_Main).End(xlDown).Row
    
    For r = ItemRow + 1 To lr
    
        collItem.Add .Cells(r, c_Main).Value
        collItemRule.Add .Cells(r, c_Main).Font.ColorIndex
        
    Next
    
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

    lr = .Cells(Rows.count, 1).End(xlUp).Row

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

StartDate = obj.GetStartDate
EndDate = obj.GetEndDate

count = EndDate - StartDate + 1

With shtDiary

    If .Cells(2, 1) = StartDate Then Exit Sub

    MyRowH = 25
    
    lr = .Cells(Rows.count, 1).End(xlUp).Row
    
    For r = 2 To count + 1
        
        'If r = 110 Then
        
        'ElseIf r > 110 Then
        
        'End If
        myAddDay = .Cells(r, "J")
        .Cells(r, 1) = r - 1
        .Cells(r, 2) = StartDate + r - 2 + myAddDay
        
        .Rows(r).RowHeight = MyRowH
        
        If .Cells(r, 5) <> "" Then .Rows(r).AutoFit
        
        If .Rows(r).RowHeight < MyRowH Then .Rows(r).RowHeight = MyRowH
    
    Next

    Do Until lr <= count + 1
    
        .Rows(r).Delete
    
        lr = lr - 1
    
    Loop
    
End With


End Sub

Sub Init() '於frmData表單新增施作位置、工程項目、試體項目

With frmData

    For i = 1 To collChannel.count
        
        .cboChannel.AddItem collChannel.Item(i)
    
    Next
    
    For i = 1 To collItem.count '設定不顯示的百分比項目
    
        If collItemRule.Item(i) <> 5 Then .cboItem.AddItem collItem.Item(i)
    
    Next
    
    For i = 1 To collTest.count
    
        .cboItem.AddItem collTest.Item(i)
    
    Next

    .txtDay = Format(Now(), "yyyy/mm/dd")

End With

End Sub

Sub Init_Mix() '於MixData表單新增工程項目、試體項目

With MixData

    For i = 1 To collItem.count '設定不顯示的百分比項目
    
        If collItemRule.Item(i) <> 5 Then .cboItem.AddItem collItem.Item(i)
    
    Next
    
    For i = 1 To collTest.count
    
        .cboItem.AddItem collTest.Item(i)
    
    Next
    
End With

End Sub

Sub RetrunUnit() '查詢TargetItem對應剩餘數、單位、預儲存未扣除之數量

Call GetMainRow

TargetItem = frmData.cboItem

With shtMain

    For r = ItemRow + 1 To .Cells(Rows.count, c_Main).End(xlUp).Row
    
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

    For r = ItemRow + 1 To .Cells(Rows.count, c_Main).End(xlUp).Row
    
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

Private Sub GetMainRow()

Call ReturnMainRow(arr)

ConRow = arr(0)
TestRow = arr(1)
ChannelRow = arr(2)
ItemRow = arr(3)
PgsRow = arr(4)

End Sub



