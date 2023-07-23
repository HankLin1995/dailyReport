VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPCCES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private wbname As String
Private wb As Object
Private myFun As New clsMyfunction

Private Sub Class_Initialize()

'Application.ScreenUpdating = False

End Sub

Sub getFileName(Optional ByVal f As String) '取得預算書內容

If f = "" Then f = Application.GetOpenFilename

If f = "False" Then MsgBox "未取得檔案", vbCritical: End

tmp = Split(f, "\")

wbname = tmp(UBound(tmp))


Workbooks.Open (f)

Set wb = Workbooks(wbname)

End Sub

Private Function getShtContentName()

For Each sht In wb.Sheets

    If sht.Name Like "*詳細表" Then
    
        getShtContentName = sht.Name
    
    End If

Next

End Function

Sub getAllContents()

Call clearBudget

Dim coll As New Collection

With wb.Sheets(getShtContentName)

    '.Unprotect ("RATSWVNXYRCMPIZ") 'only read it's ok

    lr = .Cells(.Rows.count, "F").End(xlUp).Row
    
    For r = 9 To lr
    
        item_num = myFun.tranCharcter_NUM(.Cells(r, 1))
        item_cont = .Cells(r, 2)
        item_unit = .Cells(r, 3)
        item_amount = .Cells(r, 4)
        item_cost = .Cells(r, 5)
        
        If item_num <> "" Then 'And item_amount <> 0 Then
        
            item_cont = getFullName(r)

            Call getData(item_num, item_cont, item_unit, item_amount, item_cost)

            myResult = Join(Array(item_num, item_cont, item_unit, item_amount, item_cost), ":")

            Debug.Print myResult
        
        End If
    
    Next
    
End With

'MsgBox "取得PCCES預算書完成，請進行部分項目微調", vbInformation

wb.Close False

End Sub

Private Function getFullName(ByVal r As Integer)

Set wb = Workbooks(wbname)

With wb.Sheets(getShtContentName)

    If .Cells(r + 1, "F") <> "" Then
    
        getFullName = myFun.tranCharcter(.Cells(r, "B"))
        
    Else
    
        lr = .Cells(r, "F").End(xlDown).Row
        
        For rr = r To lr - 1
        
            If cnt > 0 And .Cells(rr, "A") <> "" Then Exit For
        
            s = s & myFun.tranCharcter(.Cells(rr, "B"))
            cnt = cnt + 1
        
        Next
    
        getFullName = s
        
    End If

End With

End Function

Private Sub getData(ByVal num As String, ByVal cont As String, ByVal unit As String, ByVal amount As String, ByVal cost As String)

    With ThisWorkbook.Sheets("Budget")
    
        lr = .Cells(.Rows.count, 1).End(xlUp).Row
        
        r_Main = lr + 1
        
        .Cells(r_Main, 1) = num
        .Cells(r_Main, 2) = cont
        .Cells(r_Main, 3) = unit
        .Cells(r_Main, 4) = amount
        .Cells(r_Main, 5) = cost
        
        .Cells(r_Main, 6) = "=D" & r_Main & "*E" & r_Main
            
        .Range("A" & r_Main & ":F" & r_Main).HorizontalAlignment = xlLeft
        .Range("A" & r_Main & ":F" & r_Main).Borders.LineStyle = xlContinuous
        .Range("D" & r_Main & ":F" & r_Main).NumberFormatLocal = "0.00"
        .Range("D" & r_Main & ":E" & r_Main).Interior.ColorIndex = 40 '鎖定儲存格用
        .Range("D" & r_Main & ":F" & r_Main).Font.Name = "Times New Roman"

    
    End With

End Sub

Sub clearBudget()

With ThisWorkbook.Sheets("Budget")

    Set rng = .Cells.SpecialCells(xlCellTypeLastCell)
    
    Debug.Print rng.Address
    
    .Range("A3:" & rng.Address).Clear
    
    '.Rows("1").ClearContents
    
    .Range("H1:" & rng.Address).Clear
    
End With

End Sub

Sub clearMainChanges()

With ThisWorkbook.Sheets("Main")

    Set rng = .Cells.SpecialCells(xlCellTypeLastCell)
    
    Debug.Print rng.Address
    
    .Range("K1:" & rng.Address).Clear
    
    .Range("B6:C6").ClearContents
    
End With

End Sub

Sub clearPAY_EX()

With ThisWorkbook.Sheets("PAY_EX")

    Set rng = .Cells.SpecialCells(xlCellTypeLastCell)
    
    Debug.Print rng.Address
    
    .Range("A2:" & rng.Address).Clear
    
End With

End Sub

Sub settingColorRules()

Set coll_2nd = getCollSeconedName

For i = 1 To coll_2nd.count

    p = p & i & "." & coll_2nd(i) & vbNewLine

Next

On Error GoTo ERRORHANDLE

targetCnt = CInt(InputBox("請輸入之後為百分比項目之序位" & vbNewLine & p, , 3))

With Sheets("Budget")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row
    
    For r = 3 To lr
    
        item_num = .Cells(r, 1)
        tmp = Split(item_num, ".")
        item_cont = .Cells(r, 2)
        item_amount = .Cells(r, 5)
        
        If UBound(tmp) + 1 = 2 Then second_cnt = second_cnt + 1
        If UBound(tmp) + 1 = 1 Then first_cnt = first_cnt + 1

        If item_amount <> 0 Then
        
            .Cells(r, 7) = 3
            
            If second_cnt >= targetCnt Then .Cells(r, 7) = 5
            If first_cnt > 1 Then .Cells(r, 7) = 7
        
        End If
        
        If item_cont Like "*工程保險*" Then .Cells(r, 7) = 3
    
    Next

End With

Exit Sub

Call ReArrangeTitle

ERRORHANDLE: MsgBox "等等需要使用者自行設定準則!", vbInformation

End Sub

Function getSumMoney(Optional ByVal contract_cnt As Integer = 0)

With Sheets("Budget")

lr = .Cells(.Rows.count, 1).End(xlUp).Row

'Set rng = .Rows("2").Find("契約數量")

Select Case contract_cnt

Case 0: c = 4
Case Else: c = 8 + (contract_cnt - 1) * 3

End Select

For r = 3 To lr
    
    item_rule = .Cells(r, 1).Font.ColorIndex
    amount = .Cells(r, c)
    price = .Cells(r, c + 1)
    note = .Cells(r, c + 2)
    
    If item_rule = 3 Or item_rule = 5 Then getSumMoney = getSumMoney + amount * price

Next

End With

End Function

Sub getPercentageItems()

Set collSeconedName = getCollSeconedName

'Debug.Print "========百分比項目========"

For Each it In collSeconedName

    sum_cost = 0

    With Sheets("Budget")
    
        lr = .Cells(.Rows.count, 1).End(xlUp).Row
    
        For r = 3 To lr
        
            item_rule = .Cells(r, 1).Font.ColorIndex
        
            If item_rule = 5 Then
            
                item_num = .Cells(r, 1)
                item_name = .Cells(r, 2)
                item_sum_cost = .Cells(r, 6)
            
                tmp = Split(item_num, ".")
                
                If UBound(tmp) >= 1 Then
                
                    If collSeconedName(Join(Array(tmp(0), tmp(1)), ".")) = it Then
                    
                        sum_cost = sum_cost + item_sum_cost
                    
                    End If
                
                End If
            
            End If
        
        Next
        
        If sum_cost <> 0 Then
            'mark 2nd itemName font colorindex
            
            Set rng = .Columns("B").Find(it)
            rng.Font.ColorIndex = 5
            
        End If
        
    End With

Next

End Sub

Sub getPercentageItems_export(ByVal t_change As Integer)

Set collSeconedName = getCollSeconedName

'Debug.Print "========百分比項目========"

c = t_change_to_column(t_change)

For Each it In collSeconedName

    sum_cost = 0

    With Sheets("Budget")
    
        lr = .Cells(.Rows.count, 1).End(xlUp).Row
    
        For r = 3 To lr
        
            item_rule = .Cells(r, 1).Font.ColorIndex
        
            If item_rule = 5 Then
            
                item_num = .Cells(r, 1)
                item_name = .Cells(r, 2)
                item_sum_cost = .Cells(r, c + 2)
            
                tmp = Split(item_num, ".")
                
                If UBound(tmp) >= 1 Then
                
                    If collSeconedName(Join(Array(tmp(0), tmp(1)), ".")) = it Then
                    
                        sum_cost = sum_cost + item_sum_cost
                    
                    End If
                
                End If
            
            End If
        
        Next
        
        If sum_cost <> 0 Then
           ' Debug.Print it & ">" & sum_cost
            arr = Array(it, "式", "1", sum_cost, sum_cost)
            Call exportToMain(arr, 5, t_change)
            'mark 2nd itemName font colorindex
            
            Set rng = .Columns("B").Find(it)
            rng.Font.ColorIndex = 5
            
        End If
        
    End With

Next

End Sub

Function getCollSeconedName()

Dim coll As New Collection

With Sheets("Budget")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row

    For r = 3 To lr
    
        item_num = .Cells(r, 1)
        item_name = .Cells(r, 2)
        
        tmp = Split(item_num, ".")
        
        If UBound(tmp) = 1 Then
        
            coll.Add item_name, item_num
        
        End If
    
    Next

End With

Set getCollSeconedName = coll

End Function

Function getRecordingItemsByRecDate(ByVal rec_date As Date) ', Optional myRules = 3)

Dim Inf_obj As New clsInformation
Dim coll As New Collection

t_change = Inf_obj.getContractChangesByDate(rec_date)

c = t_change_to_column(t_change)

With Sheets("Budget")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row

    For r = 3 To lr
    
        item_rule = .Cells(r, 1).Font.ColorIndex
    
        If item_rule = 3 Then
        
            item_num = .Cells(r, 1)
            item_name = .Cells(r, 2)
            item_amount = .Cells(r, c)

            If item_amount <> 0 Then coll.Add item_name

        End If
    
    Next

End With

Set getRecordingItemsByRecDate = coll

End Function

Function getAllItemsByRecDate(ByVal rec_date As Date, Optional myRules = 3)

Dim Inf_obj As New clsInformation
Dim coll As New Collection

t_change = Inf_obj.getContractChangesByDate(rec_date)

c = t_change_to_column(t_change)

With Sheets("Budget")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row

    For r = 3 To lr
    
        item_rule = .Cells(r, 1).Font.ColorIndex
    
        If item_rule = 3 Or item_rule = 5 Then
        
            item_num = .Cells(r, 1)
            item_name = .Cells(r, 2)
            item_amount = .Cells(r, c)

            If item_amount <> 0 Then coll.Add item_name

        End If
    
    Next

End With

Set getAllItemsByRecDate = coll

End Function

Function getPercentageItemsByColorIndex()

Dim coll As New Collection

With Sheets("Budget")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row

    For r = 3 To lr
    
        item_rule = .Cells(r, 2).Font.ColorIndex
    
        If item_rule = 5 Then
        
            item_name = .Cells(r, 2)
            coll.Add item_name

        End If
    
    Next

End With

Set getPercentageItemsByRecDate = coll

End Function

Function getRecordingItemsAll()

Dim Inf_obj As New clsInformation
Dim coll As New Collection

With Sheets("Budget")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row

    For r = 3 To lr
    
        item_rule = .Cells(r, 1).Font.ColorIndex
    
        If item_rule = 3 Then
        
            item_num = .Cells(r, 1)
            item_name = .Cells(r, 2)
            'item_amount = .Cells(r, 4)

            'If item_amount <> 0 Then
            coll.Add item_name

        End If
    
    Next

End With

Set getRecordingItemsAll = coll

End Function

Sub getRecordingItems_export(ByVal t_change As Integer)

c = t_change_to_column(t_change)

With Sheets("Budget")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row

    For r = 3 To lr
    
        item_rule = .Cells(r, 1).Font.ColorIndex
    
        If item_rule = 3 Then
        
            item_num = .Cells(r, 1)
            item_name = .Cells(r, 2)
            item_unit = .Cells(r, 3)
            item_amount = .Cells(r, c)
            item_cost = .Cells(r, c + 1)
            item_sum_cost = .Cells(r, c + 2)
        
            arr = Array(item_name, item_unit, item_amount, item_cost, item_sum_cost)
        
            Call exportToMain(arr, 3, t_change)
        
        End If
    
    Next

End With

End Sub

Sub exportToMain(arr, myColorindex, t_change)

With Sheets("Main")

    r = .Cells(.Rows.count, 6 + t_change * 5).End(xlUp).Row
    
    .Cells(r + 1, 6 + t_change * 5).Resize(1, 5) = arr
    .Cells(r + 1, 6 + t_change * 5).Resize(1, 5).Borders.LineStyle = 1
    .Cells(r + 1, 6 + t_change * 5).Font.ColorIndex = myColorindex
    
'    .Range("F" & r + 1 & ":J" & r + 1) = arr
'    .Range("F" & r + 1 & ":J" & r + 1).Borders.LineStyle = 1
'    .Range("F" & r + 1).Font.ColorIndex = myColorindex
    
End With

End Sub

Sub ReArrangeTitle() '重整準則

With Sheets("Budget")

For r = 3 To .Cells(.Rows.count, 1).End(xlUp).Row

    .Cells(r, 1).Font.ColorIndex = .Cells(r, 7)

Next

End With

End Sub

Sub clearOldReport(ByVal t_change) 'clear舊有Main工作表報表資料

With Sheets("Main")

    lr = .Cells(.Rows.count, 6 + t_change * 5).End(xlUp).Row
    
    .Cells(3, 6 + t_change * 5).Resize(lr, 5).Clear

End With

End Sub

Sub markTitle()

With Sheets("Budget")
    
    beginDate = InputBox("請輸入開工日期", , Format(Now(), "yyyy/mm/dd"))
    .Rows("1").ClearContents
    .Range("D1") = "第0次變更>" & beginDate

End With

Sheets("Main").Range("B4") = beginDate

End Sub

Sub check_item_name_repeat()

Dim coll_err As New Collection

With Sheets("Budget")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row
    
    For r = 3 To lr
    
        key_value = .Cells(r, 2)
    
        Dim myFunc As New clsMyfunction
        
        Set coll_rows = myFunc.getRowsByUser("Budget", "B", key_value)
    
        If coll_rows.count > 1 Then
            
            For i = 2 To coll_rows.count
            
                err_prompt = "第" & coll_rows(i) & "列:" & key_value & "已被使用!"
            
                On Error Resume Next
                coll_err.Add err_prompt, err_prompt
                On Error GoTo 0
            
            Next
    
        End If
    
    Next

    For i = 1 To coll_err.count
    
        prompt = prompt & coll_err(i) & vbNewLine
    
    Next

    If prompt <> "" Then MsgBox prompt, vbCritical: End

End With

End Sub

'Function getMoneyByItemKey(ByVal item_key)
'
'With Sheets("契約詳細表")
'
'r = .Columns("A").Find(item_key).Row
'c = .Rows("1").Find("數量").Column
'
'getMoneyByItemKey = .Cells(r, c + 1)
'
'End With
'
'End Function

'Function getAllItems()
'
'Dim coll As New Collection
'
'With Sheets("契約詳細表")
'
'lr = .Cells(.Rows.count, 1).End(xlUp).Row
'
'For r = 2 To lr
'
'    Item = .Cells(r, 1)
'    note = .Cells(r, "G")
'    If note = "" Then coll.Add Item, Item
'
'Next
'
'End With
'
'Set getAllItems = coll
'
'End Function

'Sub checkIsRepeat()
'
'Dim coll As New Collection
'
'With Sheets("契約詳細表")
'
'lr = .Cells(.Rows.count, 1).End(xlUp).Row
'
'For r = 2 To lr
'
'    myKey = .Cells(r, 1)
'    On Error GoTo ERRORHANDLE
'    coll.Add myKey, myKey
'
'Next
'
'If s <> "" Then
'    MsgBox s, vbCritical
'    .Range("A2").Resize(lr, 7).Clear
'Else
'    Debug.Print "契約詳細表皆為唯一值!"
'End If
'
'Exit Sub
'
'End With
'
'ERRORHANDLE:       ' 錯誤處理用的程式碼
'  s = s & "列數=" & r & ",項目=" & myKey & "已經被使用!請修正!" & vbNewLine
'  Resume Next
'
'End Sub

'Function checkCollection(ByRef coll, ByVal coll_key As String)
'
'checkCollection = coll(coll_key)
'
'End Function

Function t_change_to_column(ByVal t_change)

Select Case t_change

Case 0: t_change_to_column = 4
Case Else: t_change_to_column = 8 + (t_change - 1) * 3

End Select

End Function

Function get2ndIndex(ByVal item_index As String)

tmp = Split(item_index, ".")

ch = ""
For j = LBound(tmp) To 1
    ch = ch & "." & tmp(j)
Next

get2ndIndex = mid(ch, 2)

End Function


Private Sub Class_Terminate()

'Application.ScreenUpdating = True

End Sub
