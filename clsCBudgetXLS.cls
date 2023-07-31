VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCBudgetXLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'also need "clsBudgetItems" ClassModule

Private shtBudget As Worksheet
Private shtCBudget As Worksheet
Private obj As New clsBudgetItems
Private PCCES_obj As New clsPCCES
Private Inf_obj As New clsInformation

Private collAddress As New Collection
Private ItemCountArray As Variant
Public IsFixItemCount As Boolean
Private mode As String
Private ratio As Double

Private Sub Class_Initialize()

obj.CBudget_changeIndex

Set shtBudget = Sheets("Budget")
Set shtCBudget = Sheets("CBudget")

'取得預算書的資料
'get budget data from sheets("Budget")

Set coll_changes = Inf_obj.getContractChanges

t_change = coll_changes.Count - 1

c = PCCES_obj.t_change_to_column(t_change - 1)

'Debug.Assert c = 8

obj.getData (c)
obj.sumEachSepPrice

a = obj.sumSpecificPrice(1, 2)
b = obj.sumSpecificCPrice(1, 2)

ratio = b / a 'get the ratio to check other item

End Sub

Sub getPrintPage()

'&P=當前頁數
'&N=總頁數
'2020.05.05 update

Dim sht As Worksheet

hp = ActiveSheet.HPageBreaks.Count + 1 '縱向頁數
VP = ActiveSheet.VPageBreaks.Count + 1 '橫向頁數
PG = hp * VP '總頁數

Set sht = Sheets("CBudget")

If PG > 1 Then

sht.PageSetup.RightHeader = "第 &P+1 頁,共 &N+1頁"

sht.Range("L3") = "第1頁,共" & PG + 1 & "頁"
sht.Range("L3").Font.ColorIndex = 2

Else

sht.PageSetup.RightHeader = ""
sht.Range("L3").Font.ColorIndex = 1

End If

End Sub

Sub getMode()

mode = 1 'InputBox("1.專案工程" & vbctrl & "2.自籌款", , "1")

If mode = "1" Then

    ItemCountArray = Array("一", "二", "三", "四", "五a", "五b", "六", "七", "八", "貳", "參", "肆", "伍", "陸")

ElseIf mode = "2" Then '自籌款沒有工程保險費

    ItemCountArray = Array("一", "二", "三", "四", "五a", "五b", "六", "七", "貳", "參", "肆", "伍", "陸")

End If

End Sub

Sub ClearAll2() '清理格式

With shtCBudget

Do Until .Cells(7, 2) = ""

    .Rows(7).Delete

Loop

End With

End Sub

Sub RetriveData() '從變更預算明細表上的資料名稱去回傳應該是正確的變更預算資料

With shtCBudget

    lr = .Cells(.Rows.Count, 2).End(xlUp).Row
    
    For r = 7 To lr

        myunit = .Cells(r, 3)
        
        '若為顯示總表的時候會出現錯誤(可以為顯示明細表時解除該bug)
        '因為累計項目的col=3不為""
        'EX:主體工程工作費、雜項工程工作費...etc
        '2020.05.05 update
        
        If myunit <> "" Then
        
            myitem = .Cells(r, 2)
        
            myIndex = obj.collItemIndex(myitem)

            Call WriteData(myIndex, r)
        
        End If
        
        
    Next

End With

End Sub

Sub ReadData() 'Main function

r = 7

With shtCBudget

    For Each mSepEach In obj.collSepEach
    
        IsSingle = CheckIsSingle(mSepEach) '檢查是否為單一項目
    
        i = 0 'Controll mSep
        Count = 0
        
        Call AddTitle(r, mSepEach, IsSingle) '新增項次以及項目名稱
        
        For Each mSep In obj.collSep
        
            i = i + 1
            
            If mSep = mSepEach Then
                
                Count = Count + 1
                
                If IsSingle Then
                    Call WriteData_Single(i, r, Count)
                Else
                    Call WriteData_Mix(i, r, Count)
                End If
        
            End If
        
        Next
        
        Call AddSum(r, mSepEach, IsSingle)
    
    Next

End With

End Sub

Sub WriteData(ByVal i As Integer, ByVal r As Integer)

Item = obj.collitem(i)
unit = obj.collUnit(i)
num = obj.collNum(i)
Num_ch = obj.collCNum(i)
money = obj.collPrice(i)
totalmoney = num * money
totalmoney_ch = obj.collCNum(i) * obj.collCPrice(i)

With shtCBudget

.Cells(r, 2) = Item
.Cells(r, 3) = unit '.Cells(r, 3)
.Cells(r, 4) = num '.Cells(r, 5)
.Cells(r, 5) = Num_ch '.Cells(r, 4)

Call getDiff(r, num, Num_ch, "F", "G")

.Cells(r, "H") = money '.Cells(r, 6)
.Cells(r, "I") = totalmoney '.Cells(r, 8)
.Cells(r, "J") = totalmoney_ch '.Cells(r, 9)

Call getDiff(r, totalmoney, totalmoney_ch, "K", "L")

End With

End Sub

Sub WriteData_Single(ByVal i As Integer, ByRef r, ByVal Count As Integer)

r = r - 1

Call WriteData(i, r)
Call FormulaRow(r, 3)

shtCBudget.Cells(r, "M") = "=getSumDiff(K" & r & ",L" & r & ")"

collAddress.Add r

r = r + 1

End Sub

Sub WriteData_Mix(ByVal i As Integer, ByRef r, ByVal Count As Integer)

shtCBudget.Cells(r, 1) = Count

Call WriteData(i, r)
Call FormulaRow(r, 2)

r = r + 1

End Sub

Sub FormulaRow(ByVal r As Integer, ByVal mode As Byte)

If IsFixItemCount = True Then Exit Sub

With shtCBudget.Range("A" & r & ":M" & r)
    
    .VerticalAlignment = xlCenter
    .Borders.LineStyle = 1
    .Font.Name = "標楷體"
    .Parent.Range("H" & r & ":J" & r).NumberFormatLocal = "#,##0"
    .Parent.Range("K" & r & ",L" & r).NumberFormatLocal = "#,##"
    
    If mode = 1 Then 'for Title
    
        .Font.Size = 14
        .RowHeight = 30
    ElseIf mode = 2 Then 'for Content
    
        .Font.Size = 12
        .RowHeight = 25
        
        If Len(.Parent.Range("B" & r)) > 10 Then
        
            .Parent.Range("B" & r).WrapText = True
            .EntireRow.AutoFit
            If .RowHeight < 25 Then .RowHeight = 25
    
        End If
    
    ElseIf mode = 3 Then 'for Content(only one)
    
        .Font.Size = 14
        .RowHeight = 30
    
    End If

End With

End Sub

Sub AddSum(ByRef r, ByVal SepName As String, ByVal IsSingle As Boolean)

Call FormulaRow(r, 1)

If IsSingle Then Exit Sub

With shtCBudget

    collAddress.Add r

    .Cells(r, 2) = SepName & "計"

    r = r + 1
    
End With

End Sub

Sub AddTitle(ByRef r, ByVal SepName As String, ByVal IsSingle As Boolean)

On Error GoTo ERRORHANDLE

Call FormulaRow(r, 1)

Count = obj.collSepEachIndex(SepName)

shtCBudget.Cells(r, 1) = "(" & ItemCountArray(Count - 1) & ")"
shtCBudget.Cells(r, 2) = SepName

'if count=

r = r + 1

Exit Sub

ERRORHANDLE:
End

End Sub

Sub ChangeCellColor() '改變變更後的文字顏色

With shtCBudget

    lr = .Cells(.Rows.Count, 2).End(xlUp).Row
    
    For r = 7 To lr
    
        If .Cells(r, "E") <> .Cells(r, "D") Then .Cells(r, "E").Font.ColorIndex = 3
        If .Cells(r, "I") <> .Cells(r, "J") Then .Cells(r, "J").Font.ColorIndex = 3
        
        .Range("F" & r & ":G" & r & ",K" & r & ":M" & r).Font.ColorIndex = 3
        
    Next
    
    .PageSetup.PrintArea = "A1:M" & lr '2020.05.05 update

End With

End Sub

Sub useSumFormula()

With shtCBudget

lr = .Cells(.Rows.Count, 2).End(xlUp).Row

For r = 7 To lr

    If .Cells(r, 1) = "1" Then sr = r
    If .Cells(r, 1) = "" Then
        er = r
        
        .Cells(r, "I") = "=SUM(I" & sr & ":I" & er - 1 & ")"
        .Cells(r, "J") = "=SUM(J" & sr & ":J" & er - 1 & ")"
        .Cells(r, "K") = "=SUM(K" & sr & ":K" & er - 1 & ")"
        .Cells(r, "L") = "=SUM(L" & sr & ":L" & er - 1 & ")"
        
        Sum = .Cells(r, "K") '正項目相加
        CSum = .Cells(r, "L") '負項目相加
        '.Cells(r, "M") = "=getSumDiff(K" & r & ",L" & r & ")" ' getSumDiff(Sum, CSum)"
        .Cells(r, "M") = getSumDiff(Sum, CSum)
    End If
    
Next

End With

End Sub

Sub DealSpecificSum()

With shtCBudget

    lr = .Cells(.Rows.Count, 2).End(xlUp).Row

    For r = lr To 3 Step -1
        
        Select Case .Cells(r, 1)
    
        Case "(一)"
            .Cells(r, 1).EntireRow.Insert xlShiftUp, xlFormatFromRightOrBelow
            .Cells(r, 1) = "(壹)"
            .Cells(r, 2) = "發包工作費"
    
        Case "(三)"
        
            .Cells(r, 1).EntireRow.Insert xlShiftUp
            .Cells(r, 2) = "(一)~(二)小計"
            Call loopSumFormula(r - 1, 2)
            .Cells(r, "M") = "=getSumDiff(K" & r & ",L" & r & ")"
            '.Cells(r, "M") = getSumDiff(.Cells(r, "K"), .Cells(r, "L"))

        Case "(六)"
        
            .Cells(r, 1).EntireRow.Insert xlShiftUp
            .Cells(r, 2) = "(五a)+(五b)小計"
            Call loopSumFormula(r - 1, 6, 5)
            .Cells(r, "M") = "=getSumDiff(K" & r & ",L" & r & ")"
        
        Case "(七)"
            
            If mode = "2" Then
            
                .Cells(r + 1, 1).EntireRow.Insert xlShiftUp
                .Cells(r + 1, 2) = "發包工作費總額"
                Call loopSumFormula(r, 7 + 1)
                .Cells(r + 1, "M") = "=getSumDiff(K" & r + 1 & ",L" & r + 1 & ")"
    
                '.Cells(r + 1, 1).EntireRow.Insert xlShiftUp
                '.Cells(r + 1, 2) = "(三)~(七)小計"
                'Call loopSumFormula(r, 7 + 1, 3)
                '.Cells(r + 1, "M") = "=getSumDiff(K" & r + 1 & ",L" & r + 1 & ")"

            End If

        Case "(八)"
        
            If mode = "1" Then
            
                .Cells(r + 1, 1).EntireRow.Insert xlShiftUp
                .Cells(r + 1, 2) = "發包工作費總額"
                Call loopSumFormula(r, 8 + 1)
                .Cells(r + 1, "M") = "=getSumDiff(K" & r + 1 & ",L" & r + 1 & ")"
          
                '.Cells(r + 1, 1).EntireRow.Insert xlShiftUp
                '.Cells(r + 1, 2) = "(三)~(八)小計"
                'Call loopSumFormula(r, 8 + 1, 3)
                '.Cells(r + 1, "M") = "=getSumDiff(K" & r + 1 & ",L" & r + 1 & ")"

            End If
        
        End Select
            
        If r = lr Then
        
            .Cells(r + 1, 1).EntireRow.Insert xlShiftUp
            .Cells(r + 1, 2) = "總工程費"
            Call loopSumFormula(r, collAddress.Count)
            .Cells(r + 1, "M") = "=getSumDiff(K" & r + 1 & ",L" & r + 1 & ")"
            '.Cells(r + 1, "M") = getSumDiff(.Cells(r + 1, "K"), .Cells(r + 1, "L"))
            
        End If
    
    Next

End With

End Sub

Private Function CheckIsSingle(ByVal mSep As String)

With shtBudget
    
    Set brng = .Columns(1).Find(what:=mSep)
    
    Set arng = .Columns(1).FindNext(brng)
    
    If arng.Address = brng.Address Then
        CheckIsSingle = True
    Else
        CheckIsSingle = False
    End If

End With

End Function

Private Function getDiff(ByVal r As Integer, ByVal before As Double, ByVal after As Double, _
                        ByVal addcol As String, ByVal minuscol As String)

With shtCBudget

    If before > after Then
        .Cells(r, minuscol) = before - after
    ElseIf before < after Then
        .Cells(r, addcol) = after - before
    End If

End With

End Function

Private Sub loopSumFormula(ByVal r As Integer, ByVal Count As Integer, Optional ByVal scount As Integer = 1)

arr = Array("I", "J", "K", "L")

For i = 0 To UBound(arr)

    shtCBudget.Cells(r + 1, arr(i)) = getSumFormula(Count, arr(i), scount)

Next

End Sub

Private Function getSumFormula(ByVal Count As Integer, ByVal col As String, Optional scount As Integer = 1)

For i = scount To Count

    f = f & "+" & col & collAddress(i)

    getSumFormula = "=SUM(" & mid(f, 2) & ")"

Next

End Function

Sub getAllReport(ByVal IsHidden As Boolean)

arr = Array("I", "J", "K", "L", "M")

With shtCBudget

    lr = .Cells(.Rows.Count, 2).End(xlUp).Row
    
    For r = 7 To lr
    
        If .Cells(r, 1) Like "(*" Then
          
          targetRow = getSumRow(r, .Cells(r, 2), IsSingle)
          
          If IsSingle = False Then
          
            For i = LBound(arr) To UBound(arr)
            
              If IsHidden = True Then
                .Cells(r, arr(i)) = "=" & arr(i) & targetRow
                .Cells(r, "C") = "式"
                .Cells(r, "D") = 1
                .Cells(r, "E") = 1
              Else
                .Cells(r, arr(i)) = ""
                .Range("C" & r & ":E" & r) = ""
              End If
              
            Next
            
            .Cells(targetRow, 1).EntireRow.Hidden = IsHidden
          
          End If
          
        ElseIf .Cells(r, 1) = "" Then
        
        Else
        
                .Cells(r, 1).EntireRow.Hidden = IsHidden

        End If
    
    Next

If IsHidden = True Then

    For Each cmt In .Comments
        
        cmt.Delete

    Next cmt
    
    '當分頁強制分開時，可以將強制分開的格線打破，讓總表維持在一頁
    
    If .HPageBreaks.Count <> 0 Then
        .HPageBreaks(1).DragOff xlDown, 1
    End If
    .PageSetup.PrintArea = "A1:M" & lr

End If

End With

End Sub

Sub CheckRatio()

tmp = InputBox("目前的檢查係數為" & ratio & vbctrl & "請輸入欲檢查的分類項目:", _
                , "三,四,五a,六,七")

chtmps = Split(tmp, ",")

With shtCBudget
    lr = .Cells(.Rows.Count, 2).End(xlUp).Row
    For r = 7 To lr
        
        If .Rows(r).Hidden = False Then
        
            For Each chtmp In chtmps
                
                If .Cells(r, 1) = "(" & chtmp & ")" Then
                
                    ratioMoney = .Cells(r, "I") * ratio
                    
                    Debug.Print .Cells(r, 1) & ":" & ratioMoney
                    
                    If .Cells(r, "J") > ratioMoney Then MsgBox .Cells(r, 1) & .Cells(r, 2) & "編制過高!!", vbCritical
                    
                    On Error Resume Next
                    .Cells(r, "J").Comment.Delete
                    On Error GoTo 0

                    .Range("J" & r).AddComment
                    .Range("J" & r).Comment.Visible = False
                    .Range("J" & r).Comment.Text Text:=CStr(Round(ratioMoney, 3))
                    
                    Exit For
                    
                End If
                
            Next
        
        End If
        
    Next

    For Each cmt In .Comments
    
      With cmt.Shape.TextFrame.Characters.Font
        .Name = "Times New Roman"
        .Size = 14
      End With
      
    Next cmt
    
End With

End Sub

Function getSumRow(ByVal r As Integer, ByVal Sep As String, ByRef IsSingle)

With shtCBudget

    Set rng = .Columns.Find(what:=Sep & "計")

    If rng Is Nothing Then
        getSumRow = r
        IsSingle = True
    Else
        getSumRow = rng.Row
        IsSingle = False
    End If

End With

End Function

Private Sub Class_Terminate()

obj.CBudget_recover

End Sub
