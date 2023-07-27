Attribute VB_Name = "test"
Sub test_pastePhoto()

Dim o As New clsReportPhoto
Dim Inf_obj As New clsInformation

Sheets("ReportPhoto").Range("A1") = Inf_obj.conName

msg = MsgBox("是否列印PDF?", vbYesNo)

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
'檢試驗管制總表
'品質抽查表


Sub Test2()

a = "G:\我的雲端硬碟\ExcelVBA\監造日報表\@監造日報表DEV\施工照片\1120130-鋼板樁打設抽查\123.jpg"



initialFolder = mid(a, 1, InStrRev(a, "\"))

Debug.Print initialFolder

End Sub




