Attribute VB_Name = "NumAndSum"
Sub test_calcRatio()

With Sheets("Num")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row
    lc = .Cells(1, 1).End(xlToRight).Column
    
    For c = 2 To lc
    
        today_sum = .Cells(lr, c)
        all_sum = Sheets("Sum").Cells(lr, c)
    
        ratio = today_sum / 18419249
        
        For r = 70 To 74
        
            .Cells(r, c) = ratio
            
        Next
        
    Next

End With

End Sub

Function getTotalMoney()

With Sheets("Num")

lr = .Cells(1, 1).End(xlDown).Row

For r = 2 To lr
    
    If .Cells(r, 1).Font.ColorIndex = 3 Then
    
    End If

Next

End With

End Function



Sub test_calcSumMoney()

Set coll = getMoneyColl

With Sheets("Num")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row
    lc = .Cells(1, 1).End(xlToRight).Column
    
    For c = 2 To lc
    
        sum_money = 0
    
        For r = 2 To lr - 1
        
            item_name = .Cells(r, 1)
            item_money = coll(.Cells(r, 1))
            
            item_num = .Cells(r, c)
            
            sum_money = sum_money + item_money * item_num
            
        Next
    
        .Cells(lr, c) = sum_money
    
    Next

End With

End Sub

Sub unittest_getMoneyColl()

Set coll = getMoneyColl

item_name = "よよ"

Debug.Assert coll(item_name) = 118

End Sub

Function getMoneyColl()

Dim coll As New Collection

With Sheets("Main")

lr = .Cells(.Rows.count, "I").End(xlUp).Row

For r = 3 To lr

    item_name = .Cells(r, "F")
    item_money = .Cells(r, "I")

    coll.Add item_money, item_name

Next

End With

Set getMoneyColl = coll

End Function

Sub test_KeyInPGS()

Dim o As New clsReport

With Sheets("Report")

si = 82 ' InputBox("秨﹍计=")
ei = 104 'InputBox("挡计=")

For i = si To ei

    .Range("K2") = i

    pgs = .Range("I6")
    
    Call o.KeyInPGS(.Range("C2").Value, pgs)

Next

End With

End Sub

Sub test_changeLen()

Dim coll As New Collection

With Sheets("Records")

lr = .Cells(.Rows.count, 1).End(xlUp).Row

For r = 3 To lr

    If .Cells(r, "K") = 36.7 Then
        targetNum = .Cells(r, 1)
        On Error Resume Next
        coll.Add targetNum, targetNum
        On Error GoTo 0
    End If

Next

For Each it In coll
    For r = 3 To lr
    
        If .Cells(r, 1) = it Then
            
            rec_num = .Cells(r, "F")
            'Debug.Print "NEW=" & rec_num * 16.6 / 16.7
            .Cells(r, "F") = rec_num * 36.6 / 36.7
        End If
    Next
Next

End With

End Sub

Sub getSumNumMain()

With Sheets("Records")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row
    lc = .Cells(2, 1).End(xlToRight).Column
    
    .Range("A3:K" & lr).Sort key1:=.Range("B3:B" & lr), order1:=xlAscending '逼

End With

With Sheets("Num")

lr = .Cells(1, 1).End(xlDown).Row
lc = .Cells(1, 1).End(xlToRight).Column

.Range("B2").Resize(lr, lc).ClearContents

For r = 2 To lr

    For c = 2 To lc
    
        mydate = .Cells(1, c)
        
        item_name = .Cells(r, 1)
        
        If IsUsedItem(item_name) = True Then
        
        .Cells(r, c) = getSumNumByDateAndItemName(mydate, item_name)

        End If

    Next

Next

End With

Call test_calcHold
Call test_calcSum

MsgBox "俱ЧΘ舘~"

End Sub

Function IsUsedItem(ByVal item_name)

IsUsedItem = True

With Sheets("Records")

    Set rng = .Columns("E").Find(item_name)
    
    If rng Is Nothing Then IsUsedItem = False

End With

End Function

Function getSumNumByDateAndItemName(ByVal d As Date, ByVal item_name As String)

With Sheets("Records")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row
    lc = .Cells(2, 1).End(xlToRight).Column

    Set rng = .Columns("B").Find(d)
    
    If Not rng Is Nothing Then
    
        For r = rng.Row To lr
        
            rec_date = .Cells(r, 2)
            
            If rec_date <> d Then Exit For
            
            rec_item = .Cells(r, 5)
            
            If rec_item = item_name Then
                
                rec_num = .Cells(r, 6)
                
                getSumNumByDateAndItemName = getSumNumByDateAndItemName + rec_num

            End If
        
        Next
    
    End If

End With

End Function

Sub test_calcHold()

Dim rec_num As Double

With Sheets("Num")

lc = .Cells(1, 1).End(xlToRight).Column
lr = .Cells(1, 1).End(xlDown).Row

For r = 2 To lr

    'If r = 34 Then Stop

    item_name = .Cells(r, 1)
    
    If IsNumOnlyOne(item_name) = False Then

        hold = 0
    
        For c = 2 To lc
        
            rec_num = .Cells(r, c)
        
            'sum_num = rec_num  '仓璸计秖
            
            rec_num_final = Int(rec_num + hold) '(セら计秖+逞緇计秖)俱计场だ
            
            hold = rec_num + hold - rec_num_final '(セら计秖+逞緇计秖)计场だ
            
            If rec_num <> 0 Then
                
                'Debug.Print "rec_num=" & rec_num
                'Debug.Print "sum_num=" & rec_num_final
                'Debug.Print "hold=" & hold
                
                .Cells(r, c) = rec_num_final
                
            End If
        
        Next
        
    End If

Next

End With

End Sub

Function IsNumOnlyOne(ByVal item_name)

With Sheets("Main")

    Set rng = .Columns("F").Find(item_name)
    
    Debug.Assert Not rng Is Nothing
    
    If rng.Offset(0, 2) = 1 Then
        IsNumOnlyOne = True
    Else
        IsNumOnlyOne = False
    End If

End With

End Function

Sub test_calcSum()

Dim rec_num As Double

With Sheets("Num")

For c = .Columns.count To 1 Step -1
    lr_test = .Cells(1, c).End(xlDown).Row
    lr = .Rows.count
    If lr_test <> lr Then lc_num = c: Exit For
    
Next

lc = .Cells(1, 1).End(xlToRight).Column
lr = .Cells(1, 1).End(xlDown).Row

Sheets("Sum").Range("B2").Resize(lr, lc).ClearContents

For r = 2 To lr

    sum_num = 0
    'con_num = getContractNum(.Cells(r, 1))

    For c = 2 To lc_num
    
        rec_num = .Cells(r, c)
    
        sum_num = sum_num + rec_num '仓璸计秖
        
        'If sum_num > con_num Then sum_num = con_num '代刚琌禬筁计秖璶ノ硂兵
            
        Sheets("Sum").Cells(r, c) = sum_num
            
    Next

Next

End With

End Sub

Function getContractNum(ByVal item_name As String)

With Sheets("Main")

    Set rng = .Columns("F").Find(item_name)
    
    getContractNum = .Cells(rng.Row, "H")

End With

End Function

Sub test_countFull()

With Sheets("Sum")

lr = .Cells(.Rows.count, 1).End(xlUp).Row

lc = .Cells(1, 1).End(xlToRight).Column

For r = 2 To lr

    tNum = getContractNum(.Cells(r, 1))

    For c = 2 To lc
    
        If .Cells(r, c) > tNum Then
            .Cells(r, c) = tNum
            Debug.Print .Cells(r, 1)
        End If
    
    Next

Next

End With

End Sub

