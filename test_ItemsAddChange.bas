Attribute VB_Name = "test_ItemsAddChange"
Function getDiffItems(ByVal shtName As String, ByVal sr As Integer, ByVal col_eng As String)

Dim coll As New Collection

With Sheets(shtName)

lr = .Cells(.Rows.count, 1).End(xlUp).Row

For r = sr To lr

    s = .Range(col_eng & r)

    On Error Resume Next
    
    coll.Add s, s
    
    On Error GoTo 0

Next

End With

Set getDiffItems = coll

End Function

Sub addNotZeroItems()

Set coll = getCollNotZero

For Each it In coll

    Debug.Print it

Next

End Sub

Function getColl_SumNum()

Dim coll_sumNum As New Collection

Set coll_recItems = getDiffItems("Records", 3, "E")

For Each RecItem In coll_recItems

    Set collRows = getFindRowsByOneCol(RecItem, "E")
    
    sumNum = 0
    
    For Each rngAddress In collRows
    
        recNum = Sheets("Records").Range(rngAddress).Offset(0, 1)
    
        sumNum = sumNum + recNum
    
    Next

    'Debug.Print RecItem & ":" & sumNum
    
    coll_sumNum.Add sumNum, RecItem

Next

Set getColl_SumNum = coll_sumNum

End Function

Function getCollNotZero()  '_compareConNum

Dim coll_not_zero_num As New Collection

Set coll_sumNum = getColl_SumNum

With Sheets("Main")

t_change = 1

c = 6 + t_change * 5

lr = .Cells(.Rows.count, c).End(xlUp).Row

For r = 3 To lr

    'If r = 23 Then Stop

    conName = .Cells(r, c)
    conRule = .Cells(r, c).Font.ColorIndex
    conNum = .Cells(r, c + 2)
    
    If conRule = 3 And conNum <> 0 Then
    
        On Error Resume Next
    
        sumNum = getSumNumByCollsumNum(conName, coll_sumNum)
        
        If Err.Number = 0 Then
        
            If Round(sumNum - conNum, 4) <> 0 Then
            
                Debug.Print conName & ":" & sumNum & "," & conNum
                
                coll_not_zero_num.Add conName
            
            End If
        Else
        
            Debug.Print conName & ":" & sumNum & "," & conNum
            
            coll_not_zero_num.Add conName
        
        'Debug.Print "ConName=" & ConName & ";conNum=" & conNum & ",sumNum=" & sumNum
        
        End If
    
    End If

Next

End With

Set getCollNotZero = coll_not_zero_num

End Function

Function getSumNumByCollsumNum(ByVal conName As String, ByVal coll_sumNum) As Double

On Error GoTo ERRORHANDLE

getSumNumByCollsumNum = coll_sumNum(conName)

ERRORHANDLE:

Exit Function

getSumNumByCollsumNum = 0

End Function

Function getFindRowsByOneCol(ByVal FindValue As String, ByVal col_eng As String)
    
    Dim coll As New Collection
    
    With Worksheets("Records").Columns(col_eng)
        Set rng = .Find(FindValue, LookIn:=xlValues)
        
        'Debug.Print rng.Address
        coll.Add rng.Address
        
        Set rng_next = .FindNext(rng)
        
        Do Until rng_next.Address = rng.Address
        
            'Debug.Print rng_next.Address
            coll.Add rng_next.Address
            
            Set rng_next = .FindNext(rng_next)
        
        Loop
        
    End With
    
    Set getFindRowsByOneCol = coll
    
End Function
