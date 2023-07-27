VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private shtRecord As Object
Private shtMain As Object
Private shtDiary As Object

Private recDate As Date
Private RecChannelName As String
Private RecLocation As String
Private RecItem As String
Private RecAmount As Double
Private RecUnit As String
Private RecLast As Double
Private RecCheck As New Collection
Private RecCheckType As New Collection
Private RecDateArr As Variant

Private myFunc As New clsMyfunction

Private Sub Class_Initialize()

Set shtRecord = Sheets("Records")
Set shtMain = Sheets("Main")
Set shtDiary = Sheets("Diary")

End Sub

Sub ReadData() 'ByRef IsLocOK As Boolean)

'IsLocOK = True

With frmData

recDate = .txtDay
RecChannelName = .cboChannel
RecLocation = .txtWhere
RecItem = .cboItem
If .txtAmount <> "" Then RecAmount = .txtAmount
RecUnit = .lblUnit

'If myFunc.IsNumericWithPlusAndParentheses(RecLocation) = True Then
'
'    For Each my_loc In myFunc.SplitAllLocs(RecLocation)
'
'        err_msg = getMixLocPrompt_REC(RecItem, my_loc, RecChannelName)
'
'        If err_msg <> "" Then p1 = p1 & err_msg & vbNewLine
'
'    Next
'
'    If p1 <> "" Then MsgBox p1, vbCritical: IsLocOK = False
'
'Else
'    MsgBox RecLocation & "格式不對，無法檢查!", vbInformation
'    IsLocOK = True
'End If

'CheckTable = mid(.txtCheckTable, 2)
'
'tmp = Split(CheckTable, ",")
'
'For i = 0 To UBound(tmp, 1) Step 2
'
'    Select Case tmp(0 + i)
'
'    Case "開挖(放樣)工程查驗表": MyCheck = "開挖"
'    Case "土方工程施工查驗表": MyCheck = "土方"
'    Case "鋼筋施工組立作業查驗表": MyCheck = "鋼筋"
'    Case "模板施工組立作業查驗表": MyCheck = "模板"
'    Case "混凝土施工作業查驗表": MyCheck = "混凝土"
'    Case "碎石級配工程施工查驗表": MyCheck = "碎石級配"
'    Case "瀝青混凝土工程施工查驗表": MyCheck = "瀝青混凝土"
'    Case "傾倒式閘門工程施工查驗表": MyCheck = "傾倒式閘門"
'    Case "吊門機製造及安裝工程施工查驗表": MyCheck = "吊門機"
'    Case "自動測報工程施工查驗表": MyCheck = "自動測報"
'    Case "電氣工程施工查驗表": MyCheck = "電氣"
'    Case "焊接工程施工查驗表": MyCheck = "焊接"
'
'    End Select
'
'    RecCheck.Add MyCheck
'
'    Select Case tmp(1 + i)
'
'    Case "查驗停留點": MyCheckType = 1
'    Case "施工抽查點": MyCheckType = 0
'
'    End Select
'
'    RecCheckType.Add MyCheckType

'Next

End With

End Sub

'Function IsRecLocInvolved(Optional ByVal mode As Byte = 1)
'
'IsRecLocInvolved = False
'
''RecLocation = .txtWhere
''RecItem = .cboItem
'
'If RecLocation Like "*~*" Then
'
'    tmp = Split(RecLocation, "、")
'
'    For i = LBound(tmp) To UBound(tmp)
'
'        loc_tmp = Split(tmp(i), "~")
'
'        For j = LBound(loc_tmp) To UBound(loc_tmp)
'
'            my_loc = loc_tmp(j)
'
'            If mode = 1 Then '組合工項
'
'            p1 = p1 & getMixLocPrompt("", RecItem, my_loc) & vbNewLine
'
'            Else '單一工項
'
'            p1 = p1 & getMixLocPrompt(RecItem, "", my_loc) & vbNewLine
'
'            End If
'
'        Next
'
'    Next
'
'    If p1 <> "" Then MsgBox p1, vbCritical: IsRecLocInvolved = True
'
'End If
'
'End Function


Sub ReadData_Mix() 'ByRef IsLocOK As Boolean)

'IsLocOK = True

Dim myFunc As New clsMyfunction

With MixData_Main

    recDate = .txtDay
    RecChannelName = .cboChannel
    RecLocation = .txtWhere
    
'    If myFunc.IsNumericWithPlusAndParentheses(RecLocation) = True Then
'
'        RecItem = .cboItem
'        If .txtAmount <> "" Then RecAmount = .txtAmount
'
'        For Each my_loc In myFunc.SplitAllLocs(RecLocation)
'
'            If myFunc.IsNumericWithPlusAndParentheses(CStr(my_loc)) = True Then
'
'                err_msg = getMixLocPrompt_MIX(RecItem, my_loc)
'
'                If err_msg <> "" Then p1 = p1 & err_msg & vbNewLine
'
'            End If
'
'        Next
'
'        If p1 <> "" Then MsgBox p1, vbCritical: IsLocOK = False
'
'    Else
'
'        MsgBox RecLocation & "格式不對，無法檢查!", vbInformation
'        IsLocOK = True
'
'    End If
    
End With

End Sub

Sub Recording()

Dim obj As New clsReportTest

'IsOK = IsRecording(frmData) '排除於非工期及負數

'If IsOK = False Then Exit Sub

With shtRecord

'    For i = 1 To RecCheck.count
'
'        Check = Check & "," & RecCheck.Item(i)
'        CheckType = CheckType & "," & RecCheckType.Item(i)
'
'    Next
'
'    Check = mid(Check, 2)
'    CheckType = mid(CheckType, 2)

    

    lr = .Cells(Rows.Count, 1).End(xlUp).Row + 1

    'arr = Array( lr - 2, RecDate, RecChannelName, RecLocation, RecItem, RecAmount, RecUnit, Check, CheckType)
    arr = Array("B-" & lr - 2, recDate, RecChannelName, RecLocation, RecItem, RecAmount, RecUnit) ', Check, CheckType) '20211209
    
'    For c = 1 To UBound(arr) + 1
'
'        .Cells(lr, c) = arr(c - 1)
'
'    Next
    
    Call myFunc.AppendData("Records", arr)

    'Call obj.FindTestNum(RecItem, RecAmount, RecLast, recDate)

End With

End Sub

Sub Recording_Mix()

'IsOK = IsRecording(MixData_Main) '排除於非工期及負數

'If IsOK = False Then Exit Sub

With shtRecord

'    For i = 1 To RecCheck.count
'
'        Check = Check & "," & RecCheck.Item(i)
'        CheckType = CheckType & "," & RecCheckType.Item(i)
'
'    Next
'
'    Check = mid(Check, 2)
'    CheckType = mid(CheckType, 2)

    Dim obj As New clsMixData
    
    obj.ReadData
    Set collMix = obj.getCollMix
    
    MixArr = collMix(RecItem)
    MixName = RecItem
    MixAmount = RecAmount

    lr = .Cells(Rows.Count, 1).End(xlUp).Row + 1

    For i = 1 To UBound(MixArr)

        RecItem = MixArr(i, 1)
        RecAmount = MixArr(i, 2) * MixAmount
        RecUnit = MixArr(i, 3)

        If i = 1 Then
            arr = Array("M-" & lr - 2, recDate, RecChannelName, RecLocation, RecItem, RecAmount, RecUnit, , , MixName, MixAmount)
        Else
            arr = Array("M-" & lr - 2, recDate, RecChannelName, RecLocation, RecItem, RecAmount, RecUnit)
        End If
        
        Call myFunc.AppendData("Records", arr)

    Next

End With

End Sub

Sub Recording_Detail() 'useless

With shtRecord

    DetailArr = Split(frmData.txtDetailTable.Text, ",")

    lr = .Cells(Rows.Count, 1).End(xlUp).Row + 1

    For i = 1 To UBound(DetailArr) Step 3

        RecItem = DetailArr(i)
        RecAmount = DetailArr(i + 1)
        RecUnit = DetailArr(i + 2)

        If i = 1 Then
            arr = Array("D-" & lr - 2, recDate, RecChannelName, RecLocation, RecItem, RecAmount, RecUnit, Check, CheckType, MixName, MixAmount)
        Else
            arr = Array("D-" & lr - 2, recDate, RecChannelName, RecLocation, RecItem, RecAmount, RecUnit)
        End If
    
        For c = 1 To UBound(arr) + 1
        
            .Cells(lr + j, c) = arr(c - 1)
        
        Next
        j = j + 1

    Next

End With

End Sub

Private Function IsRecording(ByVal frm As UserForm)

IsRecording = True

Dim inf As New clsInformation

startDate = inf.GetStartDate
endDate = inf.GetEndDate

With frm

    RecDay = CDate(.txtDay)

    If RecDay < startDate Or RecDay > endDate Then
    
        IsRecording = False
        MsgBox "紀錄日期為:" & startDate & "~" & endDate, vbCritical
    
    End If
    
    '判斷是否為查驗表使用?
    If .txtCheckTable <> "" And .lblLast = "" Then Exit Function
    On Error GoTo ERRORHANDLE
    RecLast = .lblLast
    
    If RecLast < 0 Then
        
        IsRecording = False
        MsgBox "剩餘數量已變成負數!", vbCritical

    End If

End With

Exit Function

ERRORHANDLE: MsgBox "沒有選擇工項", vbCritical: End

End Function

'Sub CollectRecDate() '取得Records工作表中的日期與其第一筆對應欄位
'
'Dim CollRecDate As New Collection
'Dim collr As New Collection
'
'With shtRecord
'
'    CollRecDate.Add .Cells(3, 2).Value
'    collr.Add 3
'
'    lr = .Cells(Rows.count, 1).End(xlUp).Row
'
'    shtRecord.Range("A3:K" & lr).Sort key1:=.Range("B3:B" & lr), order1:=xlAscending
'
'    For r = 3 To lr
'
'        TargetRecDate = .Cells(r, 2)
'
'        IsCollected = False
'
'        For i = 1 To CollRecDate.count
'
'            ComRecDate = CollRecDate.Item(i)
'
'            If ComRecDate = TargetRecDate Then
'                IsCollected = True
'                Exit For
'            End If
'
'        Next
'
'        If IsCollected = False Then
'            CollRecDate.Add TargetRecDate
'            collr.Add r
'        End If
'    Next
'
'End With
'
'ReDim RecDateArr(1 To CollRecDate.count, 1 To 2)
'
'For i = 1 To CollRecDate.count
'
'    RecDateArr(i, 1) = CollRecDate(i)
'    RecDateArr(i, 2) = collr(i)
'
'Next
'
'End Sub

Sub cmdExportToDiary_Main()

Call sortRecords

Dim myFunc As New clsMyfunction

Set coll_recDates = myFunc.getUniqueItems("Records", 3, "B")

For Each recDate_str In coll_recDates
    
    ItemAndAmount = getRecordsByDate(CDate(recDate_str))
    LocAndCheck = "" 'ready
    
    Call FindDate(CDate(recDate_str), ItemAndAmount, LocAndCheck)

Next

End Sub


''        MyRecCheck = Split(.Cells(r, 8), ",")
''        MyRecCheckImp = Split(.Cells(r, 9), ",")
''
''        If UBound(MyRecCheck) <> -1 Then
''
''            For j = 0 To UBound(MyRecCheck)
''
''                MyRecCheckImp_ch = "施工抽查點"
''                If MyRecCheckImp(j) = "1" Then MyRecCheckImp_ch = "檢驗停留點"
''
''                If RecLoc = "" Then
''                    LocAndCheck = LocAndCheck & "、[" & RecCanal & "]" & "─" & MyRecCheck(j) & "," & MyRecCheckImp_ch
''                Else
''                    If RecCanal <> "" Then
''                    LocAndCheck = LocAndCheck & "、[" & RecCanal & "," & RecLoc & "]" & "─" & MyRecCheck(j) & "," & MyRecCheckImp_ch
''                    Else
''                    LocAndCheck = LocAndCheck & "、[" & RecCanal & "" & RecLoc & "]" & "─" & MyRecCheck(j) & "," & MyRecCheckImp_ch
''                    End If
''                End If
''
''            Next
''
''        End If
'
'    Next


Function getRecordsByDate(ByVal recDate As Date) 'for diary

Dim myFunc As New clsMyfunction

Set coll_rows = myFunc.getRowsByUser("Records", "B", recDate)

For i = 1 To coll_rows.Count

    r = coll_rows(i)
    
    With Sheets("Records")

        If mid(.Cells(r, 1), 1, 1) = "M" And .Cells(r, "J") <> "" Then
        
            If .Cells(r, 4) = "" Then
            s = s & "," & .Cells(r, 3) & ":" & .Cells(r, "J") & "=" & .Cells(r, "K") & getDetailUnitByMixName(.Cells(r, "J")) ' " 單位"
                   
            Else
            s = s & "," & .Cells(r, 3) & "[" & .Cells(r, 4) & "]:" & .Cells(r, "J") & "=" & .Cells(r, "K") & getDetailUnitByMixName(.Cells(r, "J")) ' " 單位"
            
            End If
        
        ElseIf mid(.Cells(r, 1), 1, 1) = "B" And .Cells(r, "F") <> 0 Then
            
            If .Cells(r, 4) = "" Then
        
            s = s & "," & .Cells(r, 3) & ":" & .Cells(r, "E") & "=" & .Cells(r, "F") & .Cells(r, "G") 'getDetailUnitByMixName(.Cells(r, "J")) ' " 單位"
    
            Else
            s = s & "," & .Cells(r, 3) & "[" & .Cells(r, 4) & "]:" & .Cells(r, "E") & "=" & .Cells(r, "F") & .Cells(r, "G") 'getDetailUnitByMixName(.Cells(r, "J")) ' " 單位"
       
            End If
    
        End If
    
    End With

Next

getRecordsByDate = mid(s, 2)

End Function

Function getDetailUnitByMixName(ByVal mix_name As String)

With Sheets("Mix")

Set rng = .Columns("A").Find(mix_name)

Debug.Assert Not rng Is Nothing

getDetailUnitByMixName = .Cells(rng.Row, 8)

End With

End Function

Private Sub FindDate(ByVal recDate As Date, ByVal s1 As String, ByVal s2 As String)

With shtDiary
    
    key_value = Format(recDate, "yyyy/mm/dd(aaa)")
    Set rng = .Columns("B").Find(what:=key_value, LookIn:=xlValues)
    
    r = rng.Row
        
    If s1 <> "" And .Cells(r, "E").Interior.ColorIndex = -4142 Then .Cells(r, 5) = s1
    If s2 <> "" And .Cells(r, "F").Interior.ColorIndex = -4142 Then .Cells(r, 6) = s2
            

End With

End Sub

Sub getTmpData() 'ByVal IsBlank As Boolean)

Dim myFunc As New clsMyfunction
Dim PCCES_obj As New clsPCCES

Set shtTmp = Sheets("TMP")

With shtTmp

    Set rng = .Cells.SpecialCells(xlCellTypeLastCell)
    .Range("A2:" & rng.Address).Clear
    
End With

With shtRecord
    
    Set coll_chname = myFunc.getUniqueItems("Records", 3, , "渠道名稱")
    Set coll_items = PCCES_obj.getRecordingItemsAll

    For Each mychname In coll_chname

        For Each item_name In coll_items
        
            Set coll_rows = myFunc.getRowsByUser2("Records", mychname, 2, "渠道名稱")
    
            rec_sum = 0
    
            For Each r In coll_rows
            
                rec_item = .Cells(r, 5)
                rec_num = .Cells(r, 6)
                rec_unit = .Cells(r, 7)
                
                If rec_item = item_name Then
                
                    rec_sum = rec_sum + rec_num
                
                End If
            
            Next
            
            arr = Array(mychname, item_name, rec_sum, rec_unit)
            Call myFunc.AppendData("TMP", arr)
        
        Next

    Next
    
    shtTmp.Columns.AutoFit
    
End With
    
End Sub

'        rr = 1
'
'        For r = 3 To lr
'
'            If .Cells(r, 3) = mychname Then 'putin basic data
'
'                shtTmp.Cells(rr, 1 + c) = .Cells(r, 3)
'                shtTmp.Cells(rr, 2 + c) = .Cells(r, 5)
'                shtTmp.Cells(rr, 3 + c) = .Cells(r, 6)
'
'                rr = rr + 1
'
'            End If
'
'        Next
'
'        With shtTmp
'
'            .Range("A1:C" & rr).Sort key1:=.Range("B1:B" & rr), order1:=xlAscending
'            '.Range(Chr(64 + c + 1) & "1:" & Chr(64 + c + 3) & rr).Sort key1:=.Range(Chr(64 + c + 2) & "1:" & Chr(64 + c + 2) & rr), order1:=xlAscending
'
'             For r = 1 To rr - 1 'collect base item
'
'                 Item = .Cells(r, 2)
'                 On Error Resume Next: coll_item.Add Item, Item: On Error GoTo 0
'
'             Next
'
'             .Cells(1, 4 + cc) = .Cells(1, 1) 'title
'
'             item_r = 2
'
'             For Each myitem In coll_item
'
'                 Sum = 0
'                 For r = 1 To rr - 1
'
'                     If .Cells(r, 2) = myitem Then
'
'                         Sum = Sum + .Cells(r, 3)
'
'                     End If
'
'                 Next
'
'                If Sum <> 0 Then
'
'                 .Cells(item_r, 4 + cc) = myitem
'                 .Cells(item_r, 5 + cc) = Sum
'
'                End If
'
'                 item_r = item_r + 1
'
'             Next
'
'             If IsBlank = False Then Set coll_item = Nothing
'
'             cc = cc + 2
'
'        End With
'
'    Next
'
'End With
'
'shtTmp.Columns("A:C").Delete Shift:=xlToLeft
'shtTmp.Columns.AutoFit
'
'End Sub

Sub Tmp2TmpTotal()

Set shtTmp = Sheets("TMP")
Set shtTmpTotal = Sheets("TMPTOTAL")
Dim colltmp As New Collection

With shtTmp

For c = 1 To .Cells(1, .Columns.Count).End(xlToLeft).Column Step 2

    chname = .Cells(1, c)

    lr = .Cells(.Rows.Count, c).End(xlUp).Row

    For r = 2 To lr

        myitem = .Cells(r, c)
        myAmount = .Cells(r, c + 1)

        If myitem <> "" Then colltmp.Add r, myitem

    Next

    With shtTmpTotal

        For total_c = 1 To .Cells(1, .Columns.Count).End(xlToLeft).Column

            total_chname = .Cells(1, total_c)

            If total_chname = chname Then

                total_lr = .Cells(.Rows.Count, 5).End(xlUp).Row

                For total_r = 2 To total_lr

                    total_item = .Cells(total_r, 5)

                    On Error Resume Next

                    lastRow = myRow
                    myRow = colltmp(total_item)

                    If lastRow <> myRow Then
                        myAmount = shtTmp.Cells(myRow, c + 1)
                        .Cells(total_r, total_c + 1) = myAmount
                    End If
                Next

                Exit For

            End If

        Next

    End With

Next

End With

End Sub
'

Sub ChangeMixToRecord(ByVal mix_name As String)

'If mix_name = "版橋Type1" Then Stop

With Sheets("Records")

Set coll = getMixCollByMixName(mix_name)

For Each mix_number In coll

    Call getSrEr(mix_number, sr, er)
    
    amount = .Cells(sr, "K")
    
    cnt = er - sr + 1
    
    For r = sr To er
        
        rec_item = .Cells(r, "E")
        rec_amount = .Cells(r, "F")
        mix_amount = getMixAmount(mix_name, rec_item) * amount
        
        If rec_amount <> mix_amount Then
        
            'Debug.Print "施作位置=" & .Cells(sr, "D") & ",列=" & r & "項目=" & rec_item & ",數值由" & rec_amount & ">>" & mix_amount
            Debug.Print mix_name & ">" & r & ">" & rec_item
            Debug.Print "rec=" & rec_amount & ",mix=" & mix_amount
            
            p = p & "第" & r & "列:" & mix_name & ">" & rec_item & " 由" & rec_amount & "改為" & mix_amount & vbNewLine
            
            .Cells(r, "F") = mix_amount
            .Cells(r, "F").Font.ColorIndex = 3
            
        End If
        
    Next
    
Next

End With

If p <> "" Then MsgBox "本次工項更正明細:" & vbNewLine & p, vbInformation

End Sub

Function getMixAmount(ByVal mix_name As String, ByVal mix_item As String)

With Sheets("Mix")

Set rng = .Columns("A").Find(mix_name)

sr = rng.Row
lr = .Cells(.Rows.Count, "D").End(xlUp).Row

For r = sr To lr

    If .Cells(r, "D") = mix_item Then getMixAmount = .Cells(r, "E"): Exit Function

Next


End With

End Function

Sub getSrEr(ByVal mix_number As String, ByRef sr, ByRef er)

With Sheets("Records")

Set rng = .Columns.Find(mix_number)

If rng Is Nothing Then MsgBox mix_number & "在record工作表找不到!", vbCritical: End

sr = rng.Row

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = sr To lr

    If .Cells(r, 1) <> mix_number Then er = r - 1: Exit Sub

Next

End With

End Sub

Function getMixCollByMixName(ByVal target_mix_name As String)

Dim coll As New Collection

With Sheets("Records")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 3 To lr

    mix_name = .Cells(r, "J")
    mix_number = .Cells(r, "A")

    If mix_name = target_mix_name Then
    
        coll.Add mix_number
    
    End If

Next

Set getMixCollByMixName = coll

End With

End Function

Sub getNumAndSumByItemName(ByVal item_name As String, ByVal report_date, ByRef RecAmount, ByRef RecSumAmount)

Dim myFunc As New clsMyfunction

Set rec_rows = myFunc.getRowsByUser("Records", "E", item_name)

RecAmount = 0
RecSumAmount = 0

For i = 1 To rec_rows.Count

    r = rec_rows(i)
    
    rec_date = CDate(Sheets("Records").Cells(r, 2))
    
    ItemAmount = Sheets("Records").Cells(r, "F")
    
    If rec_date = report_date Then RecAmount = RecAmount + ItemAmount
    If rec_date <= report_date Then RecSumAmount = RecSumAmount + ItemAmount

Next

End Sub

Sub getReportSum()

Sheets("Report_Sum").Range("A9:L40").ClearContents

Dim REC_obj As New clsRecord
Dim myFunc As New clsMyfunction
Dim PCCES_obj As New clsPCCES
Dim Inf_obj As New clsInformation

Set coll_items = PCCES_obj.getRecordingItemsAll

myFunc.showList (coll_items)

s_date = CDate(InputBox("請輸入統計起始日期", , Inf_obj.startDate))
e_date = CDate(InputBox("請輸入統計結束日期", , Format(Now(), "yyyy/mm/dd")))

Sheets("Report_Sum").Range("A3") = s_date
Sheets("Report_Sum").Range("D3") = e_date

t_change = Inf_obj.getContractChangesByDate(e_date)

For Each rec_item In coll_items

    cnt = cnt + 1

    Call REC_obj.getNumAndSumByItemName(rec_item, s_date, rec_num_s, rec_sum_s)
    Call REC_obj.getNumAndSumByItemName(rec_item, e_date, rec_num_e, rec_sum_e)

    Set coll_rows = myFunc.getRowsByUser2("Budget", rec_item, 2, "工程項目")
    
    item_unit = Sheets("Budget").Cells(coll_rows(1), 3)
    item_num = Sheets("Budget").Cells(coll_rows(1), PCCES_obj.t_change_to_column(t_change))
    report_num = rec_sum_e - rec_sum_s
    
    Call exportReport_Sum(rec_item, item_unit, item_num, report_num, rec_sum_e, cnt)
    
Next

End Sub

Sub exportReport_Sum(ByVal rec_item As String, ByVal item_unit As String, ByVal item_num, _
    ByVal report_num, ByVal rec_sum_e, ByVal cnt As Integer)

With Sheets("Report_Sum")

c = 1

If cnt > 31 Then c = 7

arr = Array(rec_item, , item_unit, item_num, report_num, rec_sum_e)

.Cells(9 + cnt - 1, c).Resize(1, 6) = arr

End With

End Sub

Sub sortRecords()

With shtRecord

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

.Range("A3:K" & lr).Sort key1:=shtRecord.Range("B3:B" & lr), order1:=xlAscending

End With

End Sub

'==========function=============

'TODO:
'

Function getMixLocPrompt_REC()
'ByVal rec_name As String, ByVal rec_loc As String, ByVal rec_canal) ', ByVal loc_style As String)

rec_name = frmData.cboItem
rec_loc_all = frmData.txtWhere
rec_canal = frmData.cboChannel

Set coll_rec_locs = myFunc.SplitAllLocs(rec_loc_all)
Set coll_rows = myFunc.getRowsByUser2("Records", rec_name, 2, "施作項目")

With Sheets("Records")

    For Each r In coll_rows

        item_loc_origin = .Cells(r, "D")
        my_rec_canal = CStr(.Cells(r, "C"))
        
        If my_rec_canal = rec_canal And myFunc.IsNumericWithPlusAndParentheses(CStr(rec_loc)) = True Then
        
            Set coll_item_locs = myFunc.SplitAllLocs(item_loc_origin)
            
            For Each rec_loc In coll_rec_locs
            
                For Each item_loc In coll_item_locs
                
                    If myFunc.IsRecLocPass(rec_loc, item_loc) = False Then
                        getMixLocPrompt_REC = "第" & r & "列:【" & item_loc & "】與表單填報【" & rec_loc & "】衝突!": Exit Function
                    End If
                
                Next
            
            Next
            
        End If
    
    Next

End With

End Function

Function getMixLocPrompt_MIX() 'ByVal mix_name As String, ByVal rec_loc As String) ', ByVal loc_style As String)

mix_name = MixData_Main.cboItem
rec_loc_all = MixData_Main.txtWhere

Set coll_rec_locs = myFunc.SplitAllLocs(rec_loc_all)
Set coll_rows = myFunc.getRowsByUser2("Records", mix_name, 2, "組合工項")

With Sheets("Records")

    For Each r In coll_rows
    
        item_loc_origin = .Cells(r, "D")

        If myFunc.IsNumericWithPlusAndParentheses(CStr(rec_loc)) = True Then
        
            Set coll_item_locs = myFunc.SplitAllLocs(item_loc_origin)
        
            For Each rec_loc In coll_rec_locs
        
                For Each item_loc In coll_item_locs
            
                    If myFunc.IsRecLocPass(rec_loc, item_loc) = False Then
                        getMixLocPrompt_MIX = "第" & r & "列:【" & item_loc & "】與表單填報【" & rec_loc & "】衝突!": Exit Function
                    End If
                
                Next
            
            Next
            
        End If
        
    Next

End With

End Function

Function getExistLocByRecDate(ByVal check_date) ' As String)

Dim coll_locs As New Collection
Dim myFunc As New clsMyfunction

Set coll_rows = myFunc.getRowsByUser("Records", "B", check_date)

'Set coll_locs = myFunc.getUniqueItemsInCollRows("Records", coll_rows, 2, , "施作位置")

With Sheets("Records")

    For Each r In coll_rows
    
        rec_canal = .Cells(r, "C")
        rec_loc = .Cells(r, "D")
    
        rec_loc_all = rec_canal & "," & rec_loc
        
        On Error Resume Next
        
        If rec_loc <> "" Then coll_locs.Add rec_loc_all, rec_loc_all
        
        On Error GoTo 0
    
    Next
    
    For Each rec_loc In coll_locs
    
        j = j + 1
        p = p & j & "." & rec_loc & vbNewLine
    
    Next
    
    If p <> "" Then i = InputBox("請選擇【" & check_date & "】曾經施作位置:" & vbNewLine & p)

End With

On Error GoTo ERRORHANDLE

getExistLocByRecDate = coll_locs(CInt(i))

Exit Function

ERRORHANDLE:

End Function
