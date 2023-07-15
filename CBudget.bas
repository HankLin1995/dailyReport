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

Sub addNewChangeItems()

Dim o As New clsInformation

Set coll = o.getContractChanges

With Sheets("Budget")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row
    lc = .Cells(2, .Columns.count).End(xlToLeft).Column
    
    .Range("D2:F" & lr).Copy .Cells(2, lc + 1)
    cnt = coll.count  'InputBox("請輸入本次為第幾次變更設計", , 1)
    changeDate = InputBox("請輸入變更設計日期", , Format(Now(), "yyyy/mm/dd"))
    .Cells(1, lc + 1) = "第" & cnt & "次變更" & ">" & CDate(changeDate)
    .Cells(1, lc + 1).Font.ColorIndex = 3
    .Cells(1, lc + 1).Resize(1, 3).Merge
    .Cells(1, lc + 1).Resize(1, 3).EntireColumn.AutoFit
    
End With

End Sub

'Sub test_budgetStored()
'
'Dim o As New clsBudgetDB
'
'If o.IsExisted("B", "契約") Then
'
'    msg = MsgBox("已經存有契約資料,是否覆蓋?", vbYesNo)
'
'    If msg = vbYes Then
'
'        Call o.clearRows("B", "契約")
'
'    Else
'
'        MsgBox "動作已結束!": Exit Sub
'
'    End If
'
'End If
'
'With Sheets("Budget")
'
'    lr = .Cells(.Rows.count, 1).End(xlUp).Row
'
'    For r = 3 To lr
'
'        item_index = .Cells(r, 1)
'        item_name = .Cells(r, 2)
'        item_unit = .Cells(r, 3)
'        item_num = .Cells(r, 4)
'        item_amount = .Cells(r, 5)
'        item_sum = .Cells(r, 6)
'
'        arr = Array("契約", item_index, item_name, item_unit, item_num, item_amount, item_sum)
'
'    '    Debug.Print UBound(arr)
'
'        o.AppendData (arr)
'
'    Next
'
'End With
'
'End Sub


