VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public pay_date As Date

Private shtPay As Object
Private PCCES_obj As New clsPCCES
Private Inf_obj As New clsInformation
Private myFunc As New clsMyfunction

Private Sub Class_Initialize()

Set shtPay = Sheets("PAY")

End Sub

Sub clearPAY()

With Sheets("PAY")

    Set rng = .Cells.SpecialCells(xlCellTypeLastCell)
    
    Debug.Print rng.Address
    
    .Range("A2:" & rng.Address).Clear

End With

End Sub

Sub getPayItems()

With Sheets("PAY")

Set coll_items = PCCES_obj.getAllItemsByRecDate(CDate(pay_date))

For Each it In coll_items

    lr = .Cells(.Rows.count, 2).End(xlUp).Row
    .Cells(lr + 1, 2) = it

Next

ThisWorkbook.Names.Add Name:="PAY_DATE", RefersTo:=pay_date

End With

End Sub

Sub getOtherInf()

Dim REC_obj As New clsRecord

t_change = Inf_obj.getContractChangesByDate(pay_date)

With Sheets("PAY")

    lr = .Cells(.Rows.count, 2).End(xlUp).Row
    
    For r = 2 To lr
    
        item_name = .Cells(r, 2)
        
        Set coll_rows = myFunc.getRowsByUser("Budget", "B", item_name)
        
        With Sheets("Budget")
        
            r_data = coll_rows(1)
        
            item_index = .Cells(r_data, 1)
            item_unit = .Cells(r_data, 3)
            c = PCCES_obj.t_change_to_column(t_change)
            item_contract_amount = .Cells(r_data, c)
            item_contract_money = .Cells(r_data, c + 1)
        
        End With
        
        .Cells(r, 1) = item_index
        .Cells(r, 3) = item_unit
        .Cells(r, 4) = item_contract_money
        .Cells(r, 5) = item_contract_amount
        
        Call REC_obj.getNumAndSumByItemName(item_name, pay_date, rec_amount, rec_sum_amount)
        
        .Cells(r, 6) = rec_sum_amount
        .Cells(r, 7) = findExPay(item_name)
        .Cells(r, 8) = findExPay_Cost(item_name)
        .Cells(r, "J") = "=IF(I" & r & "+G" & r & ">E" & r & ",""N.G."","""")"
        
        '=====set formula======
        
        .Cells(r, 1).Resize(1, 9).Borders(xlEdgeBottom).LineStyle = xlDash
        .Cells(r, 9).Interior.ColorIndex = 19
        .Cells(r, 9).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Cells(r, 10).Font.ColorIndex = 3
        
    Next

End With

End Sub

Sub storePayItems()

pay_date = CDate(mid(ThisWorkbook.Names("PAY_DATE").RefersTo, 2))

If IsPayDateLater(pay_date) = False Then MsgBox "估驗時間不能比已經估驗的還要之前!", vbCritical: End

With Sheets("PAY")
 
    lr = .Cells(.Rows.count, 1).End(xlUp).Row

    'Set coll = myFunc.getUniqueItems("PAY_EX", 2, "F")
    
    For r = 2 To lr
    
        pay_num = CDbl(.Cells(r, "I"))
    
        If pay_num <> 0 Then
        
            item_index = .Cells(r, 1)
            item_name = .Cells(r, 2)
            item_unit = .Cells(r, 3)
            item_contract_money = .Cells(r, 4)
            pay_money = item_contract_money * pay_num
        
            arr = Array(item_index, item_name, item_unit, pay_num, pay_money, pay_date)
            Call myFunc.AppendData("PAY_EX", arr)
        
        End If
        
    Next

End With

End Sub

Sub exportPayNumToReport()

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
        c1 = 0
        c2 = 0
        c3 = 0
        For r = 2 To lr
        
            item_index = .Cells(r, 1)
            item_index_2nd = PCCES_obj.get2ndIndex(item_index)

            If item_second_name = coll_second_name(item_index_2nd) Then
            
                Debug.Print item_index
                item_name = .Cells(r, 2)
                item_unit = .Cells(r, 3)
                item_contract_money = .Cells(r, 4)
                item_num = .Cells(r, 5)
                pay_num_ex = .Cells(r, 7)
                pay_cost_ex = .Cells(r, 8)
                pay_num = .Cells(r, 9)
                
                'pay_cost_ex = pay_num_ex * item_contract_money
                pay_cost = pay_num * item_contract_money
                
                pay_num_sum = pay_num_ex + pay_num
                pay_cost_sum = pay_cost_ex + pay_cost
                
                arr = Array(item_name, item_unit, item_contract_money, _
                    pay_num_ex, pay_cost_ex, pay_num, pay_cost, _
                    pay_num_sum, pay_cost_sum, item_num * item_contract_money)
                
                Call myFunc.AppendData("PAY_Report", arr)
                
                c1 = c1 + pay_cost_ex
                c2 = c2 + pay_cost
                c3 = c3 + pay_cost_sum
                
            End If
            
        Next
    
    End With

    arr_sum = Array("計", , , , c1, , c2, , c3)
    
    Call myFunc.AppendData("PAY_Report", arr_sum)

Next

Call changeFomula

Sheets("PAY_Report").Activate

End Sub

Sub clearPAY_Report()

With ThisWorkbook.Sheets("PAY_Report")

    .Range("A9:J40").ClearContents

    Set rng = .Cells.SpecialCells(xlCellTypeLastCell)
    
    .Range("A42:" & rng.Address).Clear
    
End With

End Sub

Private Sub changeFomula()

Application.ScreenUpdating = False

With Sheets("PAY_Report")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row
    
    For r = 42 To lr
        
        If .Cells(r, 3) = "" And .Cells(r, 1) <> "計" Then
        
            Debug.Print .Cells(r, 1)
            .Cells(r, 1).Resize(1, 9).Merge
            .Cells(r, 1).Resize(1, 9).Borders.LineStyle = 1
            .Cells(r, 1).Resize(1, 9).Font.ColorIndex = 22
        Else
            .Cells(r, 1).WrapText = True
            .Cells(r, 1).Resize(1, 9).Borders.LineStyle = 1
        
        End If
        
        .Rows(r).Font.Name = "標楷體"
        .Rows(r).AutoFit
        
        If .Rows(r).RowHeight < 25 Then .Rows(r).RowHeight = 25
    
    Next

End With

Application.ScreenUpdating = True

End Sub

Sub set2ndFormula()

Dim coll As New Collection

With Sheets("PAY_Report")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row

    For r = 42 To lr
    
        If .Cells(r, 1).MergeCells = True Then
            
            coll.Add r
    
        End If
    
    Next

    coll.Add lr + 1
    
'    .PageSetup.PrintArea = "A1:I41" ' & sr & :M" & lr
    
    r_data = 9
    
    For i = 1 To coll.count - 1

        s = .Cells(coll(i), 1)
        
        sr = coll(i) + 1
        er = coll(i + 1) - 2
        
        '.PageSetup.PrintArea = "A" & sr - 1 & ":I" & er - 1
        
        'Debug.Print Join(Array(getTxtInBracket(s), sr, er), ">")
        'Debug.Print "=SUM(E" & sr & ":E" & er & ")"
        
        .Cells(r_data, 1) = getTxtInBracket(s)
        .Cells(r_data, 2) = "式"
        .Cells(r_data, 3) = "=SUM(J" & sr & ":J" & er & ")"
        .Cells(r_data, 4) = 1
        .Cells(r_data, 5) = "=SUM(E" & sr & ":E" & er & ")"
        .Cells(r_data, 6) = 1
        .Cells(r_data, 7) = "=SUM(G" & sr & ":G" & er & ")"
        .Cells(r_data, 8) = 1
        .Cells(r_data, 9) = "=SUM(I" & sr & ":I" & er & ")"
        
        r_data = r_data + 1
        
    Next
    
    .PageSetup.PrintArea = "A1:" & "I" & er + 1
    
End With

End Sub

Sub getPayInfo()

Dim myFunc As New clsMyfunction
Dim inf As New clsInformation

'Call inf.getBasicInformation

With Sheets("PAY_Report")

    .Range("A4") = "工程名稱:" & inf.conName
    .Range("A5") = "施工地點:" & inf.workLoc
    
    pay_date = CDate(mid(ThisWorkbook.Names("PAY_DATE").RefersTo, 2))
    
    'If IsPayDateLater(pay_date) = False Then msgbox End: Sheets("PAY").Activate
    
    Set coll = myFunc.getUniqueItems("PAY_EX", 2, , "估驗日期")
    
    .Range("G5") = pay_date
    .Range("D3") = getPayCounts + 1
    
    Set coll_rows = myFunc.getRowsByUser("Diary", "B", Format(pay_date, "yyyy/mm/dd(aaa)"))

    pgs = Sheets("Diary").Cells(coll_rows(1), "I")
    
    .Range("G41") = pgs

End With

End Sub

Sub fs_kill(ByVal cnt)

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.fileexists(getThisWorkbookPath & "\PAY\" & "第" & cnt & "次估驗.xls") = True Then

    Set f = fso.getFile(getThisWorkbookPath & "\PAY\" & "第" & cnt & "次估驗.xls")
    Kill f

End If

End Sub


'============function================

Function getPayCounts()

Set coll = myFunc.getUniqueItems("PAY_EX", 2, , "估驗日期")
getPayCounts = coll.count

End Function

Function getTxtInBracket(ByVal s As String)

pt = InStr(1, s, "(")

getTxtInBracket = mid(s, pt + 1, Len(s) - pt - 1)

End Function

Function IsPayDateLater(ByVal pay_date As Date)

IsPayDateLater = True

Set coll = myFunc.getUniqueItems("PAY_EX", 2, , "估驗日期")

If coll.count = 0 Then
    Exit Function
Else
    For i = 1 To coll.count
    
        pay_ex_date = CDate(coll(i))
        
        If pay_date <= pay_ex_date Then IsPayDateLater = False: Exit Function
    
    Next
End If

End Function

Function getPayDates()

Set getPayDates = myFunc.getUniqueItems("PAY_EX", 2, , "估驗日期")

End Function

Function findExPay(ByVal item_name)

With Sheets("PAY_EX")

lc = .Cells(1, 1).End(xlToRight).Column

Set coll_rows = myFunc.getRowsByUser("PAY_EX", "B", item_name)

For Each r In coll_rows
    
    pay_num = .Cells(r, 4)
    pay_sum = pay_sum + pay_num
    
Next

findExPay = pay_sum

End With

End Function

Function findExPay_Cost(ByVal item_name)

With Sheets("PAY_EX")

lc = .Cells(1, 1).End(xlToRight).Column

Set coll_rows = myFunc.getRowsByUser("PAY_EX", "B", item_name)

For Each r In coll_rows
    
    pay_cost = .Cells(r, 5)
    pay_sum_cost = pay_sum_cost + pay_cost
    
Next

findExPay_Cost = pay_sum_cost

End With

End Function


