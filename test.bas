Attribute VB_Name = "test"
Sub test_pastePhoto()

Dim o As New clsReportPhoto
Dim Inf_obj As New clsInformation

Sheets("ReportPhoto").Range("A1") = Inf_obj.conName

msg = MsgBox("�O�_�C�LPDF?", vbYesNo)

If msg = vbYes Then
    o.IsXLS = False
Else
    o.IsXLS = True
End If

If Sheets("Check").Range("E1") = "Y" Then
    o.IsShowText = True
Else
    o.IsShowText = False
End If

With Sheets("Check")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 3 To lr
    
        check_name = .Cells(r, 1)
        check_eng = .Cells(r, 2)
        check_num = .Cells(r, 3)
        
        check_photo_inf = .Cells(r, "I")
        
        If check_photo_inf <> "" Then
        
            Call o.GetReportByItem(r)
        
        End If
    
    Next
    
.Activate

End With

End Sub


'TODO:
'�˸���ި��`��
'�~���d��


Sub Test2()

a = "G:\�ڪ����ݵw��\ExcelVBA\�ʳy�����\@�ʳy�����DEV\�I�u�Ӥ�\1120130-���O�Υ��]��d\123.jpg"



initialFolder = mid(a, 1, InStrRev(a, "\"))

Debug.Print initialFolder

End Sub




