Attribute VB_Name = "tmp_code"
Sub tmp_getPoropertiesByMixName()

With Sheets("Mix")

    lr = .Cells(.Rows.Count, "D").End(xlUp).Row
    
    For r = 2 To lr
    
        prop = ""
    
        If .Cells(r, "A") Like "*����*" Then
            prop = "����"
        ElseIf .Cells(r, "A") Like "*�j��*" Then
            prop = "�j��"
        ElseIf .Cells(r, "A") Like "*����*" Then
            prop = "����"
        ElseIf .Cells(r, "A") Like "*�k��*" Then
            prop = "�k��"
        End If
    
        .Cells(r, "J") = prop
    
    Next

End With

End Sub



