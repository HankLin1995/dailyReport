Attribute VB_Name = "FunctionModel"
Sub cmdGetReportIDByDate() '20221125依日期選擇頁數

With Sheets("Report")

    mydate = .Range("C2")
    myID = .Range("K2")
    
    myNewDate = InputBox("請輸入日期，格式如" & vbNewLine & mydate, , mydate)
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

reportNum = InputBox("請輸入理應為100%的報表編號")
allowence = InputBox("請輸入校正回歸允許值", , 1)
prompt = "***校正回歸完成項目***" & vbNewLine

With Sheets("Report")

    .Range("K2") = reportNum

    Call ReportRun
    
    For r = 8 To obj.getReportLastRow
    
        conNum = .Cells(r, "F")
        sumNum = .Cells(r, "I")
        
        If conNum <> sumNum Then
        
            itemName = .Cells(r, "B")
            numDiff = Round(sumNum - conNum, 4)
            
            If Abs(numDiff) < allowence Then
            
                Call dealOverNum(itemName, numDiff)
            
                prompt = prompt & vbNewLine & itemName & ":" & numDiff
        
            End If
        
        End If
    
    Next
    
    MsgBox prompt, vbInformation

End With

End Sub

Sub dealOverNum(ByVal itemName As String, ByVal numDiff As Double) '20221122處理剩餘零星數量

With Sheets("Records")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row
    
    For r = lr To 3 Step -1
    
        recName = .Cells(r, "E")
        
        If recName = itemName Then
        
            originNum = .Cells(r, "F")
            
            adjustNum = originNum - numDiff
            
            If adjustNum > 0 Then
            
                Debug.Print itemName & ",原數量=" & originNum & ">>校正=" & adjustNum
            
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

Sub cmdgetTmpData()

Dim obj As New clsRecord

'msg = MsgBox("是否要留白?", vbYesNo)

'If msg = vbYes Then
obj.getTmpData (True)
'Else
'obj.getTmpData (False)
'End If

obj.Tmp2TmpTotal

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

Dim obj As New clsBudget

obj.CollectTitle
obj.clearOldReport
obj.ExportToReport 'should change something

End Sub

Sub cmdReArrange()

Dim obj As New clsBudget

obj.ReArrangeTitle

End Sub
Sub cmdFindBudget()

Dim obj As New clsBudget

obj.FindWorkbook
If obj.IsError = True Then Exit Sub
obj.DealBudget
obj.clearBudget
obj.CollectBudget
obj.ArrangeTitle

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

StartDate = obj.GetStartDate
EndDate = obj.GetEndDate

obj.ProgressNew

End Sub

Sub cmdExportToDiary()

Dim obj As New clsRecord

obj.CollectRecDate
obj.DealDiary
obj.GetRecDetail

End Sub
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

