Attribute VB_Name = "CBudget"
Sub cmdBudgetToCBudget() 'get basic data from worksheets("Budget")

Dim obj As New clsCBudgetXLS

Application.ScreenUpdating = False

IsClearAll = MsgBox("�O�_�n�M���즳�榡?", vbYesNo)

If IsClearAll = vbYes Then

    obj.IsFixItemCount = False
    obj.ClearAll2
    obj.getMode 'new sub to get project mode
    
Else
    obj.IsFixItemCount = True
    obj.RetriveData
    
    MsgBox "�ƶq�w�g���s��z�o�I"

    Exit Sub

End If

obj.ReadData
obj.useSumFormula
obj.DealSpecificSum
obj.ChangeCellColor
obj.getPrintPage 'new to set print page

Application.ScreenUpdating = True

MsgBox "�榡�μƶq�w�g���s���J�F�I"

End Sub

Sub getAllReport()

Dim obj As New clsCBudgetXLS

msg = MsgBox("�O�_�n����`��?", vbYesNo)

mode = InputBox("1.�ܧ�]�p" & vbNewLine & "2.�ץ��w��", , 1)

cnt = InputBox("�п�J�ĴX��(�@�B�G�B�T)", , "�@")

Application.ScreenUpdating = False

If msg = vbYes Then

    Call obj.getAllReport(True)
    Call obj.ChangeCellColor
    Call obj.CheckRatio
    Call obj.getPrintPage
    
Else

    Call obj.getAllReport(False)
    Call obj.getPrintPage
    
End If

getReportName = "���Ӫ�"
If msg = vbYes Then getReportName = "�`��"


If mode = 1 Then
    
    Sheets("CBudget").Range("A2") = "��" & cnt & "���ܧ�]�p" & getReportName

Else

    Sheets("CBudget").Range("A2") = "��" & cnt & "���ץ��w��" & getReportName

End If

Application.ScreenUpdating = True

End Sub

'Function getReportName()
'
'getReportName = "���Ӫ�"
'
'With Sheets("CBudget")
'
'For Each myRow In .Rows
'
'If myRow.Hidden = True Then
'getReportName = "�`��": Exit Function
'End If
'
'Next
'
'End With
'
'End Function

Function getSumDiff(ByVal Sum As Double, ByVal CSum As Double)

'compare sum and changesum in order to get a string to show the difference

    If CSum > Sum Then
        getSumDiff = "(-)" & Format(Abs(Sum - CSum), "#,##")
    ElseIf CSum < Sum Then
        getSumDiff = "(+)" & Format(Abs(Sum - CSum), "#,##")
    Else
        getSumDiff = ""
    End If
    
End Function

Sub cmdChangeItemsChooser()

mode = InputBox("�п�ܹw�p�n���檺�B�J" & vbNewLine & "1.�s�W�ܧ����" & _
                                            vbNewLine & "2.���ܧ���" & _
                                            vbNewLine & "3.�R���ܧ�̫����", , 1)

If mode = 1 Then

Call addNewChangeItems

ElseIf mode = 2 Then

Call editChangeDate

ElseIf mode = 3 Then

Call deleteChanges

End If

End Sub

Sub deleteChanges()

Dim Inf_obj As New clsInformation
Dim PCCES_obj As New clsPCCES

Set coll_changes = Inf_obj.getContractChanges

For Each it In coll_changes

    If j > 0 Then p = j & "." & p & it & vbNewLine
    j = j + 1
    
Next

If p <> "" Then

    msg = MsgBox("�O�_�n�R���̫�@�����ܧ�]�p?", vbYesNo)
    
    If msg = vbYes Then
    
    c = PCCES_obj.t_change_to_column(j - 1)
    
    Sheets("Budget").Cells(2, c).Resize(1, 3).EntireColumn.Delete
    
    End If

Else

    MsgBox "�d�L�ܧ�]�p���e!", vbInformation

End If

End Sub

Sub addNewChangeItems()

Dim o As New clsInformation

Set coll = o.getContractChanges

With Sheets("Budget")

    cnt = coll.Count  'InputBox("�п�J�������ĴX���ܧ�]�p", , 1)
    changeDate = InputBox("�п�J�ܧ�]�p���", , Format(Now(), "yyyy/mm/dd"))
    
    Set coll_changes = o.getContractChanges
    
    For Each it In coll_changes
    
        tmp = split(it, ">")
    
        If CDate(changeDate) <= tmp(1) Then MsgBox "��������" & tmp(1) & "�٦�!", vbCritical: End
    
    Next
    
    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    lc = .Cells(2, .Columns.Count).End(xlToLeft).Column
    
    .Range("D2:F" & lr).Copy .Cells(2, lc + 1)
    
    .Cells(1, lc + 1) = "��" & cnt & "���ܧ�" & ">" & CDate(changeDate)
    .Cells(1, lc + 1).Font.ColorIndex = 3
    .Cells(1, lc + 1).Resize(1, 3).Merge
    .Cells(1, lc + 1).Resize(1, 3).EntireColumn.AutoFit
    
    MsgBox "�ܧ�]�p���e��g������O�o�A�I��פJ����~�|�ͮ�!", vbInformation
    
End With

End Sub

Sub editChangeDate()

Dim Inf_obj As New clsInformation
Dim PCCES_obj As New clsPCCES

Set coll_changes = Inf_obj.getContractChanges

For Each it In coll_changes

    If j > 0 Then p = j & "." & p & it & vbNewLine
    j = j + 1
    
Next

If p <> "" Then i = CInt(InputBox("�w�p�n�ק諸������ĴX��?" & vbNewLine & p, , j - 1))

If i <= coll_changes.Count And i > 0 Then


    On Error GoTo ERRORHANDLE
    new_change_date = CDate(InputBox("�п�J���T���ܧ�����:", , Format(Now(), "yyyy/mm/dd")))

    new_change_str = "��" & j - 1 & "���ܧ�>" & new_change_date

    
    c = PCCES_obj.t_change_to_column(j - 1)
    
    Sheets("Budget").Cells(1, c) = new_change_str

End If

Exit Sub

ERRORHANDLE:

MsgBox "����榡�����T!...yyyy/mm/dd", vbCritical

End Sub



