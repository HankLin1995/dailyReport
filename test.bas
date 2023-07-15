Attribute VB_Name = "test"
Sub clearPAY_Report()

With ThisWorkbook.Sheets("PAY_Report")

    Set rng = .Cells.SpecialCells(xlCellTypeLastCell)
    
    .Range("A42:" & rng.Address).Clear
    
End With


End Sub

Sub test_getPayNums()

Call clearPAY_Report

Dim PCCES_obj As New clsPCCES
Dim myFunc As New clsMyfunction

Set coll_second_name = PCCES_obj.getCollSeconedName

For i = 1 To coll_second_name.count

    item_second_name = coll_second_name(i)
    
    arr_title = Array("第" & myFunc.ch(i) & "號明細表(" & item_second_name & ")")
    
    Call myFunc.AppendData("PAY_Report", arr_title)

    With Sheets("PAY")
    
        lr = .Cells(.Rows.count, 1).End(xlUp).Row
    
        For r = 2 To lr
        
            item_index = .Cells(r, 1)
            item_index_2nd = get2ndIndex(item_index)

            If item_second_name = coll_second_name(item_index_2nd) Then
            
                Debug.Print item_index
                item_name = .Cells(r, 2)
                item_unit = .Cells(r, 3)
                item_contract_money = .Cells(r, 4)
                pay_num_ex = .Cells(r, 7)
                pay_cost_ex = .Cells(r, 8)
                pay_num = .Cells(r, 9)
                
                'pay_cost_ex = pay_num_ex * item_contract_money
                pay_cost = pay_num * item_contract_money
                
                pay_num_sum = pay_num_ex + pay_num
                pay_cost_sum = pay_cost_ex + pay_cost
                
                arr = Array(item_name, item_unit, item_contract_money, pay_num_ex, pay_cost_ex, pay_num, pay_cost, pay_num_sum, pay_cost_sum)
                
                Call myFunc.AppendData("PAY_Report", arr)
                
            End If
            
        Next
    
    End With

Next

Call test_changeFomula

Sheets("PAY_Report").Activate

End Sub

Sub test_changeFomula()

Application.ScreenUpdating = False

With Sheets("PAY_Report")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row
    
    For r = 42 To lr
        
        If .Cells(r, 3) = "" Then
        
            Debug.Print .Cells(r, 1)
            .Cells(r, 1).Resize(1, 9).Merge
            .Cells(r, 1).Resize(1, 9).Borders.LineStyle = 1
            .Cells(r, 1).Resize(1, 9).Font.ColorIndex = 22
        Else
            .Cells(r, 1).WrapText = True
            .Cells(r, 1).Resize(1, 9).Borders.LineStyle = 1
        
        End If
        
        .Rows(r).AutoFit
        
        If .Rows(r).RowHeight < 25 Then .Rows(r).RowHeight = 25
    
    Next

End With

Application.ScreenUpdating = True

End Sub

Function get2ndIndex(ByVal item_index As String)

tmp = Split(item_index, ".")

ch = ""
For j = LBound(tmp) To 1
    ch = ch & "." & tmp(j)
Next

get2ndIndex = mid(ch, 2)

End Function

Sub t()

Dim o As New clsPay

o.storePayItems

Sheets("PAY_EX").Activate

End Sub

'===================================
Sub checkTestCompleted() '20230225 add

With Sheets("Test")

    lr = .Cells(.Rows.count, "C").End(xlUp).Row
    
    For r = 2 To lr
    
        TestName = .Cells(r, "A")
        calcTest = .Cells(r, "F")
        doTest = .Cells(r, "G")
        
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
