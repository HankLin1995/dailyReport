Attribute VB_Name = "CBudget"
Sub cmdBudgetToCBudget() 'get basic data from worksheets("Budget")

Dim obj As New clsCBudgetXLS

Application.ScreenUpdating = False

IsClearAll = MsgBox("是否要清除原有格式?", vbYesNo)

If IsClearAll = vbYes Then

    obj.IsFixItemCount = False
    obj.ClearAll2
    obj.getMode 'new sub to get project mode
    
Else
    obj.IsFixItemCount = True
    obj.RetriveData
    
    MsgBox "數量已經重新整理囉！"

    Exit Sub

End If

obj.ReadData
obj.useSumFormula
obj.DealSpecificSum
obj.ChangeCellColor
obj.getPrintPage 'new to set print page

Application.ScreenUpdating = True

MsgBox "格式及數量已經重新載入了！"

End Sub

Sub getAllReport()

Dim obj As New clsCBudgetXLS

msg = MsgBox("是否要顯示總表?", vbYesNo)

mode = InputBox("1.變更設計" & vbNewLine & "2.修正預算", , 1)

cnt = InputBox("請輸入第幾次(一、二、三)", , "一")

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
    
    Sheets("CBudget").Range("A2") = "第" & cnt & "次變更設計" & getReportName

Else

    Sheets("CBudget").Range("A2") = "第" & cnt & "次修正預算" & getReportName

End If

Application.ScreenUpdating = True

End Sub

Function getReportName()

getReportName = "明細表"

With Sheets("CBudget")

For Each myRow In .Rows

If myRow.Hidden = True Then
getReportName = "總表": Exit Function
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





