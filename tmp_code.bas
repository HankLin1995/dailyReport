Attribute VB_Name = "tmp_code"
Sub tmp_getPoropertiesByMixName()

With Sheets("Mix")

    lr = .Cells(.Rows.Count, "D").End(xlUp).Row
    
    For r = 2 To lr
    
        prop = ""
    
        If .Cells(r, "A") Like "*¿ûµ¬*" Then
            prop = "¿ûµ¬"
        ElseIf .Cells(r, "A") Like "*¤j©³*" Then
            prop = "¤j©³"
        ElseIf .Cells(r, "A") Like "*¥ªÀð*" Then
            prop = "¥ªÀð"
        ElseIf .Cells(r, "A") Like "*¥kÀð*" Then
            prop = "¥kÀð"
        End If
    
        .Cells(r, "J") = prop
    
    Next

End With

End Sub



