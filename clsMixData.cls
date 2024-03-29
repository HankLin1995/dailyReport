VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMixData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private shtRec As Object
Private shtMix As Object
Private shtMain As Object
Private shtMix_Sum As Object

Private collMix As New Collection
Private collChannel As New Collection

Private Sub Class_Initialize()

Set shtRec = Sheets("Records")
Set shtMix = Sheets("Mix")
Set shtMain = Sheets("Main")
'Set shtMix_Sum = Sheets("Mix_Sum")

End Sub

Function getCollMix()

Set getCollMix = collMix

End Function

Sub ReadData() '取得Mix資料collection

Dim collr As New Collection
Dim MixArr As Variant
Dim j As Integer

With shtMix

    lr = .Cells(Rows.Count, 4).End(xlUp).Row
    
    For r = 3 To lr
    
        If .Cells(r, 1) <> "" Then collr.Add r
    
    Next

    collr.Add lr + 1
    
    For i = 1 To collr.Count - 1
    
        j = 0
    
        sr = collr(i)
        er = collr(i + 1) - 1
        
        ReDim MixArr(er - sr + 1, 1 To 3)
        
        MixKey = .Cells(sr, 1)
        Mixfont = .Cells(sr, 1).Font.ColorIndex
        MixBas = .Cells(sr, 2)
        MixFin = .Cells(sr, 3)
        
        MixArr(j, 1) = MixKey
        MixArr(j, 2) = MixBas
        MixArr(j, 3) = MixFin
    
        j = j + 1
        
        For r = sr To er
        
            MixArr(j, 1) = .Cells(r, 4)
            MixArr(j, 2) = .Cells(r, 5)
            MixArr(j, 3) = .Cells(r, 6)
        
            j = j + 1
        
        Next
    
        If Mixfont = 1 Then collMix.Add MixArr, MixKey
    
    Next

End With

End Sub

Sub ReadMainData() '取得渠道collection

Call ReturnMainRow(arr)

ChannelRow = arr(2)
PgsRow = arr(4)

With shtMain
    
    For r = ChannelRow + 1 To PgsRow - 1
    
        If .Cells(r, 1) <> "" Then collChannel.Add .Cells(r, 1).Value
    
    Next

End With

End Sub

Sub Init() '初始化Mix表單
 
With MixData_Main

    For i = 1 To collMix.Count
    
        .cboItem.AddItem collMix(i)(0, 1)
    
    Next
    
    For i = 1 To collChannel.Count
        
        .cboChannel.AddItem collChannel.Item(i)
    
    Next

End With

End Sub

Sub ReturnLast() '取德目標總數

With MixData_Main

    MixKey = .cboItem
    
    If MixKey = "" Then Exit Sub
    
    arr = collMix(MixKey)
    
    .lblLast = arr(0, 3)
    .lblStore = .lblLast

End With

End Sub

Sub UsedAmount() '查詢TargetItem對應剩餘數、單位、預儲存未扣除之數量

Dim frm As UserForm
Set frm = MixData_Main

TargetItem = frm.cboItem

With shtRec

    lr = .Cells(Rows.Count, 5).End(xlUp).Row

    .Range("A3:K" & lr).Sort key1:=.Range("B3:B" & lr), order1:=xlAscending

    For r = 3 To lr
    
        MixDay = CDate(frm.txtDay)
        MixUntilDay = frm.chkUntilDay
        
        RecDay = .Cells(r, "B")
        
        If RecDay > MixDay And MixUntilDay = True Then Exit For
        
        MixItem = .Cells(r, "J")
        MixAmount = .Cells(r, "K")
    
        If .Cells(r, "J") = TargetItem Then num = num + MixAmount: 'Debug.Print "列位:" & r & "num=" & num
        
    Next

    frm.lblLast = frm.lblLast - num
    frm.lblStore = frm.lblLast

End With

End Sub

Sub CheckComplete() '確認是否已經做完

With shtMix

    Set collr = getSeperateRows()

    For i = 1 To collr.Count - 1
    
        sr = collr(i)
        er = collr(i + 1) - 1
        
        targetMix = .Cells(sr, 1)
        
        'If targetMix Like "*~*" Then MsgBox ("請避免「" & targetMix & "」含有特殊字符「~」"), vbInformation
        
        targetSum = .Cells(sr, 3)
        
        MixSum = getMixSum(targetMix)
        
        .Cells(sr, 1).Font.ColorIndex = 1
        .Cells(sr, 7) = MixSum
        
        If MixSum = targetSum Then
        
            .Cells(sr, 1).Font.ColorIndex = 3
            '.Rows(sr).Resize(er - sr + 1).Hidden = True
            
        End If
        
    Next

End With

End Sub

Sub CheckUnfoundMixName() '確認是否存在找不到名稱的MixName

With shtMix

    Set collr = getSeperateRows()

    For i = 1 To collr.Count - 1
    
        sr = collr(i)
        er = collr(i + 1) - 1
        
        targetMix = .Cells(sr, 1)
        
        'If targetMix Like "*~*" Then MsgBox ("請避免「" & targetMix & "」含有特殊字符「~」"), vbInformation
        
        Set rng = .Cells.Find(targetMix)
        
        If rng Is Nothing Then
        
             MsgBox "這個字「" & targetMix & "」搜尋找不到，請更換!"
        
        End If
        
    Next

End With

End Sub

Sub hideCompleteItem() '做好的可以隱藏起來

msg = MsgBox("是否要顯示全部?", vbYesNo)

With shtMix

    Set collr = getSeperateRows()

    For i = 1 To collr.Count - 1
    
        sr = collr(i)
        er = collr(i + 1) - 1
        
        If .Cells(sr, 1).Font.ColorIndex = 3 Then
   
            If msg = vbYes Then
                .Rows(sr).Resize(er - sr + 1).Hidden = False
            Else
                .Rows(sr).Resize(er - sr + 1).Hidden = True
            End If
            
        End If
        
    Next

End With

End Sub

Private Function getMixSum(ByVal targetMix As String) '取得已做基數

With shtRec

    lr = .Cells(Rows.Count, 1).End(xlUp).Row
    
    For r = 3 To lr
    
        If .Cells(r, "J") = targetMix Then getMixSum = getMixSum + .Cells(r, "K")
        
    Next

End With

End Function

Function getSeperateRows() '20210824 取得分割Mix資料列位們

Dim coll As New Collection

With shtMix

    .UsedRange.EntireRow.Hidden = False '摺疊會出問題!
    
    lr = .Cells(Rows.Count, 4).End(xlUp).Row
    
    For r = 3 To lr
    
        If .Cells(r, 1) <> "" Then coll.Add r
    
    Next

    coll.Add lr + 1

End With

Set getSeperateRows = coll

End Function

Sub AppendData(ByVal MixName, ByVal MixDefine, ByVal MixDefineTotal, ByVal MixUnit, ByVal MixChannel, tmp) '將表單資料放入Mix

With shtMix
    
    .UsedRange.EntireRow.Hidden = False '摺疊會出問題!
    
    lr = .Cells(Rows.Count, 4).End(xlUp).Row + 1
    
    .Cells(lr, 1) = MixName
    .Cells(lr, 2) = MixDefine
    .Cells(lr, 3) = MixDefineTotal
    .Cells(lr, 8) = MixUnit
    .Cells(lr, 9) = MixChannel
    
    For i = 0 To UBound(tmp) - 1 Step 3
    
        .Cells(lr + j, 4) = tmp(0 + i)
        .Cells(lr + j, 5) = tmp(1 + i)
        .Cells(lr + j, 6) = tmp(2 + i)
        
        j = j + 1
        
    Next

End With

End Sub

'Sub clearMixSum()
'
'With shtMix_Sum
'
'    lr = .Cells(.Rows.count, 1).End(xlUp).Row
'    .Rows(3 & ":" & lr).ClearContents
'
'End With
'
'End Sub
'
'Sub getSpecificItem(ByRef collItem, ByRef collIndex)
'
'With shtMix
'
'    lr = .Cells(.Rows.count, 4).End(xlUp).Row
'
'    For r = 3 To lr
'
'        On Error Resume Next
'
'        Item = .Cells(r, 4).Value
'        collItem.Add Item, Item
'
'        On Error GoTo 0
'
'    Next
'
'End With
'
'For i = 1 To collItem.count '取得各工項名稱不重複值的index
'
'    collIndex.Add i, collItem(i)
'
'Next
'
'End Sub
'
'Sub calcMixData(ByVal collItem, ByVal collIndex)
'
'With shtMix
'
'    Dim arr(): ReDim arr(collIndex.count - 1, 2)
'
'    Set collr = getSeperateRows
'
'    j = 3
'
'    For i = 1 To collr.count - 1
'
'        sr = collr(i)
'        er = collr(i + 1) - 1
'
'        For r = sr To er
'
'            '======各單項統計======
'
'            id = .Cells(sr, 1) '工項概要
'            Item = .Cells(r, 4) '工項名稱
'            unit = .Cells(r, 6) '單位
'            itemNum = .Cells(sr, 3) * .Cells(r, 5) '該工項總數量
'            tmp = Split(id, ",")
'
'            Sheets("Mix_Sum").Range("A" & j & ":E" & j) = Array(id, Item, itemNum, unit, tmp(0)): j = j + 1
'
'            '=======合計統計=======
'
'            arr(collIndex(Item) - 1, 0) = Item
'            arr(collIndex(Item) - 1, 1) = arr(collIndex(Item) - 1, 1) + itemNum
'            arr(collIndex(Item) - 1, 2) = unit
'
'        Next
'
'    Next
'
'    shtMix_Sum.Range("G3:I" & UBound(arr) + 3) = arr
'    shtMix_Sum.Activate
'
'End With
'
'
'End Sub

Sub clearMixSum()

With ThisWorkbook.Sheets("Mix_Tmp")

    Set rng = .Cells.SpecialCells(xlCellTypeLastCell)
    
    Debug.Print rng.Address
    
    .Range("A2:" & rng.Address).Clear
    
End With

End Sub

Sub getMixUnitNum()

Dim myFunc As New clsMyfunction

Set coll_unit_names = myFunc.getUniqueItems("Mix", 3, , "主體單元")

Call myFunc.showList(coll_unit_names)

With Sheets("Mix")

    lr = .Cells(.Rows.Count, 4).End(xlUp).Row
    
    For Each unit_name In coll_unit_names
    
        Debug.Print unit_name
    
        Set coll_rows = myFunc.getRowsByUser2("Mix", unit_name, 2, "主體單元")
        
        For Each r In coll_rows
        
            r_end = getNextNotBlankRow(r, lr)
            
            For r_data = r To r_end - 1
            
                mix_name = .Cells(r, 1)
                mix_unit_name = .Cells(r, "I")
                mix_item_name = .Cells(r_data, "D")
                mix_num = .Cells(r_data, "E")
                mix_Unit = .Cells(r_data, "F")
                
                mix_len = .Cells(r, "C")
                
                arr = Array(mix_name, mix_unit_name, mix_item_name, mix_num, mix_Unit, mix_len, mix_len * mix_num)
            
                Call myFunc.AppendData("Mix_TMP", arr)
            
            Next
        
        Next
    
    Next

End With

End Sub

Sub getMixUnitSum()

Dim myFunc As New clsMyfunction

Set coll_unit_names = myFunc.getUniqueItems("Mix", 3, , "主體單元")

Call myFunc.showList(coll_unit_names)

For Each UnitName In coll_unit_names

    Set coll_rows = myFunc.getRowsByUser2("Mix_TMP", UnitName, 1, "單元名稱")
    
    sr = coll_rows(1)
    er = coll_rows(coll_rows.Count)
    
    Set coll_uqItems = myFunc.getUniqueItemsToLastRow("Mix_TMP", sr, er, "C")
    
    With Sheets("Mix_TMP")
    
        j = 0
    
        For Each it In coll_uqItems
        
            Debug.Print UnitName
        
            unit_Sum = 0
            all_Sum = 0
        
            For r = sr To er
            
                If .Cells(r, "C") = it Then
                    mix_Unit = .Cells(r, "E")
                    mix_len = .Cells(r, "F")
                    unit_Sum = unit_Sum + .Cells(r, "D")
                    all_Sum = all_Sum + .Cells(r, "G")
                
                End If
                
            Next
            
            Debug.Print it & ">" & unit_Sum & ">" & all_Sum
            
            Call myFunc.AppendData("Mix_TMP", Array("合計", UnitName, it, unit_Sum, mix_Unit, mix_len, all_Sum))
            
            If j = 0 Then lr = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            .Range("A" & lr & ":G" & lr).Borders(xlEdgeTop).LineStyle = 1
             .Range("A" & lr & ":G" & lr).Borders(xlEdgeTop).Weight = xlMedium
            
            j = j + 1
            
        Next
    
        .Columns.AutoFit
    
    End With

Next

End Sub

Sub getMixUnitSumAll()

With Sheets("Mix_TMP")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row + 1

    .Range("A" & lr & ":G" & lr).Borders(xlEdgeTop).LineStyle = 1
    .Range("A" & lr & ":G" & lr).Borders(xlEdgeTop).Weight = xlMedium
    
    Dim myFunc As New clsMyfunction
    
    Set coll_items = myFunc.getUniqueItems("Mix_TMP", 2, , "工項名稱")
    
    For Each item_name In coll_items
    
        Set coll_rows = myFunc.getRowsByUser2("Mix_TMP", item_name, 1, "工項名稱")
        
        item_num = 0
        item_sum = 0
        
        For Each r In coll_rows
        
            item_unit = .Cells(r, "E")
        
            If .Cells(r, 1) <> "合計" Then
                
                item_num = item_num + .Cells(r, "D")
                item_sum = item_sum + .Cells(r, "G")
            
            End If
        
        Next
        
        Call myFunc.AppendData("Mix_TMP", Array("總計", , item_name, , item_unit, , item_sum))
    
    Next

End With


End Sub

Function getNextNotBlankRow(ByVal r As Integer, ByVal lr As Integer)

With Sheets("Mix")

For r_data = r + 1 To lr

    If .Cells(r_data, "A") <> "" Then getNextNotBlankRow = r_data: Exit Function
    
    If r_data = lr Then getNextNotBlankRow = lr + 1

Next

End With

End Function

Function IsMixNameUsed(ByVal MixName As String)

IsMixNameUsed = True

Dim myFunc As New clsMyfunction

Set coll_rows = myFunc.getRowsByUser2("Mix", MixName, 2, "工項概要")

If coll_rows.Count = 0 Then IsMixNameUsed = False

End Function



