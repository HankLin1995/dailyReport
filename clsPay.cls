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

Dim rec_obj As New clsRecord

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
        
        Call rec_obj.getNumAndSumByItemName(item_name, pay_date, rec_amount, rec_sum_amount)
        
        .Cells(r, 6) = rec_sum_amount
        .Cells(r, 7) = findExPay(item_name)
        .Cells(r, 8) = findExPay_Cost(item_name)
        
        '=====set formula======
        
        .Cells(r, 1).Resize(1, 9).Borders(xlEdgeBottom).LineStyle = xlDash
        .Cells(r, 9).Interior.ColorIndex = 19
        .Cells(r, 9).Borders(xlEdgeLeft).LineStyle = xlContinuous
        
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
    
        pay_num = CDbl(.Cells(r, "H"))
    
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

Function IsPayDateLater(ByVal pay_date As Date)

IsPayDateLater = True

Set coll = myFunc.getUniqueItems("PAY_EX", 2, "F")

If coll.count = 0 Then
    Exit Function
Else
    For i = 1 To coll.count
    
        pay_ex_date = CDate(coll(i))
        
        If pay_date <= pay_ex_date Then IsPayDateLater = False: Exit Function
    
    Next
End If


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


