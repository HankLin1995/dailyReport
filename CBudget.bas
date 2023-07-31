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

getReportName = "明細表"
If msg = vbYes Then getReportName = "總表"


If mode = 1 Then
    
    Sheets("CBudget").Range("A2") = "第" & cnt & "次變更設計" & getReportName

Else

    Sheets("CBudget").Range("A2") = "第" & cnt & "次修正預算" & getReportName

End If

Application.ScreenUpdating = True

End Sub

'Function getReportName()
'
'getReportName = "明細表"
'
'With Sheets("CBudget")
'
'For Each myRow In .Rows
'
'If myRow.Hidden = True Then
'getReportName = "總表": Exit Function
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

mode = InputBox("請選擇預計要執行的步驟" & vbNewLine & "1.新增變更期數" & _
                                            vbNewLine & "2.更正變更日期" & _
                                            vbNewLine & "3.刪除變更最後期數", , 1)

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

    msg = MsgBox("是否要刪除最後一期的變更設計?", vbYesNo)
    
    If msg = vbYes Then
    
    c = PCCES_obj.t_change_to_column(j - 1)
    
    Sheets("Budget").Cells(2, c).Resize(1, 3).EntireColumn.Delete
    
    End If

Else

    MsgBox "查無變更設計內容!", vbInformation

End If

End Sub

Sub addNewChangeItems()

Dim o As New clsInformation

Set coll = o.getContractChanges

With Sheets("Budget")

    cnt = coll.Count  'InputBox("請輸入本次為第幾次變更設計", , 1)
    changeDate = InputBox("請輸入變更設計日期", , Format(Now(), "yyyy/mm/dd"))
    
    Set coll_changes = o.getContractChanges
    
    For Each it In coll_changes
    
        tmp = split(it, ">")
    
        If CDate(changeDate) <= tmp(1) Then MsgBox "日期不能比" & tmp(1) & "還早!", vbCritical: End
    
    Next
    
    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    lc = .Cells(2, .Columns.Count).End(xlToLeft).Column
    
    .Range("D2:F" & lr).Copy .Cells(2, lc + 1)
    
    .Cells(1, lc + 1) = "第" & cnt & "次變更" & ">" & CDate(changeDate)
    .Cells(1, lc + 1).Font.ColorIndex = 3
    .Cells(1, lc + 1).Resize(1, 3).Merge
    .Cells(1, lc + 1).Resize(1, 3).EntireColumn.AutoFit
    
    MsgBox "變更設計內容填寫完畢後記得再點選匯入報表才會生效!", vbInformation
    
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

If p <> "" Then i = CInt(InputBox("預計要修改的日期為第幾次?" & vbNewLine & p, , j - 1))

If i <= coll_changes.Count And i > 0 Then


    On Error GoTo ERRORHANDLE
    new_change_date = CDate(InputBox("請輸入正確的變更日期為:", , Format(Now(), "yyyy/mm/dd")))

    new_change_str = "第" & j - 1 & "次變更>" & new_change_date

    
    c = PCCES_obj.t_change_to_column(j - 1)
    
    Sheets("Budget").Cells(1, c) = new_change_str

End If

Exit Sub

ERRORHANDLE:

MsgBox "日期格式不正確!...yyyy/mm/dd", vbCritical

End Sub



