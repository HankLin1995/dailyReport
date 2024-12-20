Attribute VB_Name = "許願池"
Sub getItems()

With Sheets("Mix")

lr = .Cells(.Rows.Count, 4).End(xlUp).Row

For r = 3 To lr

    s = .Cells(r, 1)
    
    If s <> "" Then
    
        Debug.Print s & ":" & .Cells(r, "I")
    
    End If

Next

End With

End Sub

Sub markSep()

Dim coll As New Collection

With Sheets("Mix-Step")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 2 To lr

    s = .Cells(r, "A") & .Cells(r, "C") & .Cells(r, "D")

    Debug.Print s

    On Error Resume Next
    
    coll.Add r, CStr(s)
    
    On Error GoTo 0

Next


For Each r In coll

.Range("A" & r & ":F" & r).Borders(xlTop).LineStyle = 1

Next

End With

End Sub

Function getExRecItem(ByVal rec_item As String)

'rec_item = "小排2-5右牆305到312"

With Sheets("Mix-Step")

Set rng = .Columns("F").Find(rec_item)

If .Cells(rng.Row, "E") > 1 Then

    For r = rng.Row To 2 Step -1
    
        If .Cells(r, "E") < .Cells(rng.Row, "E") Then
        
            'Debug.Print .Cells(r, "F")
            getExRecItem = .Cells(r, "F")
            Exit For
        
        End If
    
    Next
Else

End If

End With

End Function
