Attribute VB_Name = "test"
Sub test_rebuild_cmdExportToDiary()

Dim myFunc As New clsMyfunction

Set coll_recDates = myFunc.getUniqueItems("Records", 3, "B")

For Each recDate_Str In coll_recDates

    'Set coll_recDate_rows = myFunc.getRowsByUser("Records", "B", CDate(recDate_Str))

    Debug.Print recDate_Str
    
    Debug.Print test_getRecordsByDate(CDate(recDate_Str))
    

'    For Each r In coll_recDate_rows
'
'        Debug.Print r
'
'    Next

Next

End Sub


Function test_getRecordsByDate(ByVal recDate As Date)

Dim myFunc As New clsMyfunction
Dim o As New clsRecord

'rec_date = CDate("2023/6/5")

Set coll_rows = myFunc.getRowsByUser("Records", "B", recDate)

For i = 1 To coll_rows.count

    r = coll_rows(i)
    
    With Sheets("Records")

        If mid(.Cells(r, 1), 1, 1) = "M" And .Cells(r, "J") <> "" Then
        
            If .Cells(r, 4) = "" Then
            s = s & "," & .Cells(r, 3) & ":" & .Cells(r, "J") & "=" & .Cells(r, "K") & o.getDetailUnitByMixName(.Cells(r, "J")) ' " 單位"
                   
            Else
            s = s & "," & .Cells(r, 3) & "[" & .Cells(r, 4) & "]:" & .Cells(r, "J") & "=" & .Cells(r, "K") & o.getDetailUnitByMixName(.Cells(r, "J")) ' " 單位"
            
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

test_getRecordsByDate = mid(s, 2)

End Function

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
