Attribute VB_Name = "test"
Sub test_getUserDefinedPlotOrder()

Set collMixItems = getMixItems

myIndexs = getShowIndex(collMixItems)

End Sub

Sub checkTestCompleted() '20230225 add

With Sheets("Test")

    lr = .Cells(.Rows.count, "C").End(xlUp).Row
    
    For r = 2 To lr
    
        TestName = .Cells(r, "E")
        calcTest = .Cells(r, "G")
        doTest = .Cells(r, "H")
        
        If doTest > calcTest Then
        
            prompt = prompt & TestName & "©|§ÌØ " & doTest - calcTest & "≤’" & vbNewLine & vbNewLine
        
        End If
    
    Next

    MsgBox prompt

End With

End Sub

Sub t()
With Sheets("Num")
For c = .Columns.count To 1 Step -1
    lr_test = .Cells(1, c).End(xlDown).Row
    lr = .Rows.count
    If lr_test <> lr Then Debug.Print .Cells(1, c): Exit For
    
    'If .Cells(1, c).End(xlDown).Row <> .Rows.count Then lc = c: Exit For
Next

Debug.Print lc
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


Sub test0612() '¿À≈Á∞±Ød¬I•”Ω–≥Ê

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
        
        If .Cells(r, 4) = checkday And .Cells(r, 5) = "¿À≈Á∞±Ød¬I" Then
        
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
