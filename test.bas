Attribute VB_Name = "test"
Sub enlargeWorkDays()

Dim Inf_obj As New clsInformation
Dim myFunc As New clsMyfunction

enlargeDate = InputBox("請輸入展延開始日期", , Format(Now(), "yyyy/mm/dd"))
enlargeDays = CInt(InputBox("請輸入展延天數", , 1))

Sheets("Main").Range("B6") = Inf_obj.workDay + enlargeDays
Sheets("Main").Range("C6") = CDate(enlargeDate)

With Sheets("Diary")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row

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

End Sub


'===================================
Sub checkTestCompleted() '20230225 add

Dim REC_obj As New clsRecord
Dim Test_obj As New clsReportTest

With Sheets("Test")

    lr = .Cells(.Rows.count, "C").End(xlUp).Row
    
    For r = 2 To lr
    
        TestName = .Cells(r, "A")
        ItemName = .Cells(r, "C")
        testPeriod = .Cells(r, "D")
        
        Call REC_obj.getNumAndSumByItemName(TestName, CDate(Now()), rec_num, rec_sum)
        
        ItemName = .Cells(r, "C")
        
        Call REC_obj.getNumAndSumByItemName(ItemName, CDate(Now()), item_num, item_sum)
        
        calcTest = rec_sum
        doTest = Test_obj.getTestNeedNum(item_sum, testPeriod)
        
        If doTest > calcTest Then
        
            prompt = prompt & TestName & "尚欠缺" & doTest - calcTest & "組" & vbNewLine & vbNewLine
        
        End If
    
    Next

    If prompt <> "" Then MsgBox prompt

End With

End Sub

Function getMixSum(ByVal s As String, ByVal item_name As String)

With Sheets("Mix")

Set coll = getSpecificMixName(s)

For Each it In coll

    'Debug.Print it & ":" & sr & ">" & er
    
    Set rng = .Columns("A").Find(it)
    
    sr = rng.Row
    lr = .Cells(.Rows.count, "D").End(xlUp).Row
    er = .Cells(sr, "A").End(xlDown).Row - 1
    If er > lr Then er = lr
    
     Debug.Print it & ":" & sr & ">" & er
    
    For r = sr To er
    
        If .Cells(r, "D") = item_name Then getMixSum = getMixSum + .Cells(r, "E"): Debug.Print getMixSum
    
    Next

Next

End With

End Function

Function getSpecificMixName(ByVal like_mix_name As String)

Dim coll As New Collection

With Sheets("Mix")

lr = .Cells(.Rows.count, "D").End(xlUp).Row

For r = 3 To lr

    If .Cells(r, 1) <> "" Then
    
        mix_name = .Cells(r, 1)
    
        If mix_name Like "*" & like_mix_name & "*" Then coll.Add mix_name, mix_name
        
    
    End If

Next

End With

Set getSpecificMixName = coll

End Function


Sub test0612() '檢驗停留點申請單

Set checkdaylist = getTimeList

With Sheets("Check")

lr = .Cells(1, 1).End(xlDown).Row

For Each checkday In checkdaylist

myRow = 15

i = i + 1

With Sheets("CheckList")
 
    .Range("W4") = i
    .Range("W6") = checkday - 1
    .Cells(15, 1).Resize(10, 26).ClearContents

End With

    For r = 2 To lr
        
        If .Cells(r, 4) = checkday And .Cells(r, 5) = "檢驗停留點" Then
        
            checkitem = .Cells(r, 1)
            tmp = Split(.Cells(r, 6), ",")
            checkch = tmp(0)
            CheckLoc = tmp(1)
        
            With Sheets("CheckList")
            
                .Range("A" & myRow) = checkch
                .Range("G" & myRow) = checkday
                .Range("M" & myRow) = CheckLoc
                .Range("R" & myRow) = checkitem
            
                myRow = myRow + 1
            
            End With
        
        End If
        
    Next

    If myRow = 15 Then
        i = i - 1
    Else
        Sheets("CheckList").PrintOut
    End If

Next

End With

End Sub

Sub delCheckList()

Sheets("CheckList").Cells(15, 1).Resize(10, 26).ClearContents

End Sub

Function getTimeList()

Dim coll As New Collection

With Sheets("Check")

    lr = .Cells(1, 1).End(xlDown).Row
    
    For r = 2 To lr
        
        checkday = .Cells(r, 4)
        
        On Error Resume Next
        
        coll.Add checkday, CStr(checkday)
        
    Next

End With

Set getTimeList = coll

End Function
