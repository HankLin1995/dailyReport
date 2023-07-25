Attribute VB_Name = "test"
Sub t2()

recDate = CDate("2023/7/10")

Dim myFunc As New clsMyfunction

Set coll_rows = myFunc.getRowsByUser("Records", "B", recDate)

End Sub


Sub t()

Dim myFunc As New clsMyfunction

check_date = CDate("2023/7/10") ' Format(CDate("2023/7/10"), "yyyy/mm/dd")

Set coll_rows = myFunc.getRowsByUser("Records", "B", check_date)

Debug.Assert coll_rows.count = 9

Dim o As New clsRecord

a = o.getExistLocByRecDate(CDate("2023/7/10"))

End Sub

Private Function getCheckStyle()

msg_check_style = MsgBox("是否為檢驗停留點?", vbYesNo)

getCheckStyle = "施工抽查點"
If msg_check_style = vbYes Then getCheckStyle = "檢驗停留點"

End Function

Sub main()

Dim myFunc As New clsMyfunction

check_name = getCheckName

Call splitFileName_Check(check_name, check_name_ch, check_name_eng)

check_date = InputBox("請輸入抽查時間", , CDate(Format(Now(), "yyyy/mm/dd")))
check_style = getCheckStyle
check_loc = InputBox("請輸入地點", , "0+800左牆")
check_page = countChecks(check_name_ch) + 1

arr = Array(check_name_ch, check_name_eng, check_page, check_date, check_style, check_loc)

Call myFunc.AppendData("Check", arr)

End Sub

Function getCheckName()

Dim myFunc As New clsMyfunction

Set coll_check_names = getCheckFileNames

For i = 1 To coll_check_names.count

    p = p & i & "." & coll_check_names(i) & vbNewLine

Next

mode = InputBox("請輸入要執行的抽查表" & vbNewLine & p, , 1)

getCheckName = coll_check_names(CInt(mode))

End Function

Private Function countChecks(ByVal check_name As String)

Dim myFunc As New clsMyfunction

Set coll_rows = myFunc.getRowsByUser2("Check", check_name, 1, "查驗表(中文)")

countChecks = coll_rows.count

End Function

Function splitFileName_Check(ByVal filename As String, ByRef filename_ch, ByRef filename_eng)
    
    pt2 = InStrRev(filename, "[")

    filename_ch = mid(filename, 1, pt2 - 1)
    filename_eng = mid(filename, pt2 + 1, Len(filename) - pt2 - 1)

End Function


'===================================

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
