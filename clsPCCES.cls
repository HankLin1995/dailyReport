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
Private myfun As New clsMyfunction

Private Sub Class_Initialize()

Application.ScreenUpdating = False

End Sub

Sub getFileName(Optional ByVal F As String) '取得預算書內容

If F = "" Then F = Application.GetOpenFilename

If F = "False" Then MsgBox "未取得檔案", vbCritical: End

tmp = Split(F, "\")

wbname = tmp(UBound(tmp))

Workbooks.Open (F)

Set wb = Workbooks(wbname)

End Sub

Function getShtContentName()

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
    
        item_num = myfun.tranCharcter_NUM(.Cells(r, 1))
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
    
        getFullName = myfun.tranCharcter(.Cells(r, "B"))
        
    Else
    
        lr = .Cells(r, "F").End(xlDown).Row
        
        For rr = r To lr - 1
        
            If cnt > 0 And .Cells(rr, "A") <> "" Then Exit For
        
            s = s & myfun.tranCharcter(.Cells(rr, "B"))
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
    
End With

End Sub

Sub settingColorRules()

targetCnt = CInt(InputBox("請輸入之後為百分比項目之序位", , 3))

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

Dim myBudgetObj As New clsBudget

myBudgetObj.ReArrangeTitle

End Sub

Function getSumMoney()

With Sheets("Budget")

lr = .Cells(.Rows.count, 1).End(xlUp).Row

Set rng = .Rows("2").Find("契約數量")

c = rng.Column

For r = 3 To lr
    
    item_rule = .Cells(r, c - 3).Font.ColorIndex
    amount = .Cells(r, c)
    price = .Cells(r, c + 1)
    note = .Cells(r, c + 2)
    
    If item_rule = 3 Or item_rule = 5 Then getSumMoney = getSumMoney + amount * price

Next

End With

End Function

Function getMoneyByItemKey(ByVal item_key)

With Sheets("契約詳細表")

r = .Columns("A").Find(item_key).Row
c = .Rows("1").Find("數量").Column
    
getMoneyByItemKey = .Cells(r, c + 1)

End With

End Function

Function getAllItems()

Dim coll As New Collection

With Sheets("契約詳細表")

lr = .Cells(.Rows.count, 1).End(xlUp).Row

For r = 2 To lr
    
    Item = .Cells(r, 1)
    note = .Cells(r, "G")
    If note = "" Then coll.Add Item, Item

Next

End With

Set getAllItems = coll

End Function

Sub checkIsRepeat()

Dim coll As New Collection

With Sheets("契約詳細表")

lr = .Cells(.Rows.count, 1).End(xlUp).Row

For r = 2 To lr

    myKey = .Cells(r, 1)
    On Error GoTo ERRORHANDLE
    coll.Add myKey, myKey

Next

If s <> "" Then
    MsgBox s, vbCritical
    .Range("A2").Resize(lr, 7).Clear
Else
    Debug.Print "契約詳細表皆為唯一值!"
End If

Exit Sub

End With

ERRORHANDLE:       ' 錯誤處理用的程式碼
  s = s & "列數=" & r & ",項目=" & myKey & "已經被使用!請修正!" & vbNewLine
  Resume Next

End Sub

Function checkCollection(ByRef coll, ByVal coll_key As String)

checkCollection = coll(coll_key)

End Function

Private Sub Class_Terminate()

Application.ScreenUpdating = True

End Sub
