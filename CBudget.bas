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

If mode = 1 Then
    
    Sheets("CBudget").Range("A2") = "��" & cnt & "���ܧ�]�p" & getReportName

Else

    Sheets("CBudget").Range("A2") = "��" & cnt & "���ץ��w��" & getReportName

End If

Application.ScreenUpdating = True

End Sub

Function getReportName()

getReportName = "���Ӫ�"

With Sheets("CBudget")

For Each myRow In .Rows

If myRow.Hidden = True Then
getReportName = "�`��": Exit Function
End If

Next

End With

End Function

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





