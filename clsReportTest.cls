VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReportTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private TestRow As Integer
Private ChannelRow As Integer
Private ItemRow As Integer
Private TestArr As Variant
Private ReportDay As Date
Public t_change As Integer

Private shtMain As Object
Private shtRec As Object
Private shtBudget As Object
Private shtDiary As Object
Private shtReport As Object
Private shtTest As Object

Private Sub Class_Initialize()

Set shtMain = Sheets("Main")
Set shtRec = Sheets("Records")
Set shtBudget = Sheets("Budget")
Set shtDiary = Sheets("Diary")
Set shtReport = Sheets("Report")
Set shtTest = Sheets("Test")

ReportDay = shtReport.Range("C2")

End Sub

Sub Init_Test()

Application.EnableEvents = False

Call GetMainRow
Call Init_getTestRow(sr, er)
Call Init_DealTestRow(sr, er)

lr_test = shtMain.Cells(ChannelRow, 1).End(xlUp).Row

If lr_test < ChannelRow - 1 Then
    TestRow_End = lr_test
Else
    TestRow_End = ChannelRow - 1
End If

Count = Round((TestRow_End - TestRow) / 2, 0) - 1

Call Init_InsertTestRow(Count, sr + 3)

c = 0

For r = TestRow + 1 To TestRow_End

    ts = shtMain.Cells(r, 1)
    tsm = shtMain.Cells(r, 2)
    
    If j = Count + 2 Then c = 6: j = 0
    
    shtReport.Cells(sr + 2 + j, 2 + c) = ts
    shtReport.Cells(sr + 2 + j, 3 + c) = tsm
    
    targetAddr = shtReport.Cells(sr + 2 + j, 2 + c).Address
    
    shtMain.Cells(r, 4) = Replace(targetAddr, "$", "")
    
    j = j + 1
    
Next

Application.EnableEvents = True

End Sub

Private Sub Init_DealTestRow(ByVal sr, ByVal er)

With shtReport

For r = er - 1 To sr + 4 Step -1

    .Rows(r).Delete

Next

.Range("B" & sr + 3 & ":F" & sr + 3 & ",H" & sr + 3 & ":K" & sr + 3).ClearContents

End With

End Sub

Private Sub Init_InsertTestRow(ByVal Count As Integer, ByVal targetRow As Integer)

With shtReport

i = 0

Do Until i = Count

    .Rows(targetRow).Insert
    
    .Range("D" & targetRow & ":E" & targetRow).Merge
    
    i = i + 1
    
Loop

End With

End Sub

Private Sub Init_getTestRow(sr, er)

With shtReport

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row

    For Each rng In .Range("A1:A" & lr)
    
        If rng Like "三、*" Then
            sr = rng.Row
        ElseIf rng Like "四、*" Then
            er = rng.Row
        End If
    
    Next

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

Sub CollTestArr()

Call GetMainRow

With shtMain

    For r = TestRow + 1 To ChannelRow - 1
    
        If .Cells(r, 1) <> "" Then c = c + 1
    
    Next
    
    ReDim TestArr(1 To c, 1 To 6) '名稱,設計,$$,欄位,本日,累計
    
    i = 1
    
    For r = TestRow + 1 To TestRow + c
    
        TestName = .Cells(r, 1)
        TestDesign = .Cells(r, 2)
        
        tmp = split(TestDesign, ",")
        
        If UBound(tmp) < t_change Then
        
        TestDesign = tmp(UBound(tmp))
        
        Else
        
        TestDesign = tmp(t_change)
        
        End If
        
        'TestDesign = tmp(t_change)

        
        'If t_change = 1 Then TestDesign = .Cells(r, 5)
        
        TestMoney = .Cells(r, 3)
        TestLoc = .Cells(r, 4)
        
        arr = Array(TestName, TestDesign, TestMoney, TestLoc)
    
        For j = 0 To UBound(arr)
        
            TestArr(i, j + 1) = arr(j)
        
        Next
    
        i = i + 1

    Next

End With

Exit Sub

'TESTERRORHANDLE:

'MsgBox "請先於上的試驗組數進行修改!" & vbNewLine & "如有變更應該要多,出來", vbCritical

End Sub

Sub GetTestRec()

With shtRec

    lr = .Cells(Rows.Count, 1).End(xlUp).Row

    For i = 1 To UBound(TestArr, 1)
    
        TestName = TestArr(i, 1)
        RecAmount = 0
        RecSumAmount = 0
    
        For r = 2 To lr
        
            RecDay = .Cells(r, 2)
            ItemName = .Cells(r, 5)
            
            If ItemName = TestName Then
            
                ItemAmount = .Cells(r, 6)
                
                If RecDay = ReportDay Then RecAmount = RecAmount + ItemAmount
                    
                If RecDay <= ReportDay Then RecSumAmount = RecSumAmount + ItemAmount
        
            End If
        
        Next
        
        TestArr(i, 5) = RecAmount
        TestArr(i, 6) = RecSumAmount
        
    Next

End With

End Sub

Sub WriteReport()

With shtReport

    For i = 1 To UBound(TestArr, 1)
    
        TestDesign = TestArr(i, 2)
        TestLoc = TestArr(i, 4)
        TestAmount = TestArr(i, 5)
        TestSumAmount = TestArr(i, 6)
        col = mid(TestLoc, 1, 1)
        r = mid(TestLoc, 2)
        
        Select Case col
        
        Case "B"
        
            .Range(TestLoc).Offset(0, 1) = TestDesign
            .Range(TestLoc).Offset(0, 2) = TestAmount
            .Range(TestLoc).Offset(0, 4) = TestSumAmount
            
        Case "H"
        
            .Range(TestLoc).Offset(0, 1) = TestDesign
            .Range(TestLoc).Offset(0, 2) = TestAmount
            .Range(TestLoc).Offset(0, 3) = TestSumAmount
        
        End Select
        
    Next

End With

End Sub

'Sub FindTestNum(ByVal RecItem As String, ByVal RecAmount As Double, ByVal RecLast As Double, ByVal recDate As Date)
'
'sr = 2
'
'With shtTest
'
'    For r = sr To .Cells(.Rows.count, 1).End(xlUp).Row
'
'        targetTest = .Cells(r, 3)
'
'        If targetTest = RecItem Then
'
'            targetTestStr = .Cells(r, 5)
'            contractNum = .Cells(r, 4)
'            SumRecItemNum = contractNum - RecLast
'            oldTestItem = .Cells(r, 1)
'
'            SumTestNum = CountOldTestAmount(oldTestItem, recDate)
'
'            IsTested = CheckTestNum(SumTestNum, SumRecItemNum, targetTestStr)
'
'            Exit Sub
'
'        End If
'    Next
'
'End With
'
'End Sub

Private Function CountOldTestAmount(ByVal TestName As String, ByVal recDate As Date) As Double

With shtRec

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row

    For r = 3 To lr
    
        OldRecDate = .Cells(r, 2)
        OldItem = .Cells(r, 5)
        OldAmount = .Cells(r, 6)
        
        If recDate < OldRecDate Then Exit For
        
        If OldItem = TestName Then CountOldTestAmount = CountOldTestAmount + OldAmount
    
    Next

End With

End Function

Private Function CheckTestNum(ByVal SumTestAmount As Double, ByVal SumRecItemNum As Double, ByVal targetTestStr As String) As Boolean

tmp = split(targetTestStr, ",")

For i = LBound(tmp) To UBound(tmp)

    NeedTestItemNum = tmp(i) '200,350...etc
    
    If SumRecItemNum > Val(NeedTestItemNum) Then SumTestCount = SumTestCount + 1: j = j + 1

    

Next

If SumTestAmount < SumTestCount Then

    If j = UBound(tmp) + 1 Then

        MsgBox "施作數量" & SumRecItemNum & vbNewLine & "超過" & tmp(j - 1) & "---該做試驗了喔!!"
    Else
    
        MsgBox "施作數量" & SumRecItemNum & vbNewLine & "介於" & tmp(j - 1) & "~" & tmp(j) & "---該做試驗了喔!!"

    End If

End If

End Function

Function getTestNeedNum(ByVal num As Double, ByVal s As String)

tmp = split(s, ",")

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

