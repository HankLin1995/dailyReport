VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public IsError As Boolean

Private BudgetArr As Variant
Private wbMain As Object
Private wbBudget As Object
Private shtBudget As Object
Private shtMain As Object
Private r_Main As Integer
Private c_Main As Integer

Private color_show
Private color_percent
Private color_hide

Private collTitle As New Collection
Private collTitleRow As New Collection
Private collTitleRule As New Collection

Private Sub Class_Initialize()

Set wbMain = ThisWorkbook

Set shtBudget = wbMain.Sheets("Budget")
Set shtMain = wbMain.Sheets("Main")

IsError = False

r_Main = 3
c_Main = 6

color_show = 3
color_percent = 5
color_hide = 7

End Sub

Sub FindWorkbook() '取得預算書活頁簿

wb_path = Application.GetOpenFilename

If wb_path = "False" Then GoTo ERRHANDLE

Workbooks.Open filename:=wb_path

Set wbBudget = Workbooks(TranWbName(wb_path))

Exit Sub

ERRHANDLE:

MsgBox "尚未取得任何預算書活頁簿，請重新選取!", vbCritical

IsError = True

End Sub

Private Function TranWbName(wb_path) '從活頁簿path取得workbook.name

For i = 1 To Len(wb_path)

    If mid(wb_path, i, 1) = "\" Then Sep = i

Next

TranWbName = mid(wb_path, Sep + 1)

End Function

Sub DealBudget()

'With wbBudget '處理第二頁受保護工作表

'Sheets("預算詳細表").Select
'Sheets("預算詳細表").Copy Before:=Sheets(2)

'.Sheets(2).Cells.Copy
'.Sheets.Add .Sheets(2)
'.Sheets(2).Paste

'.Activate

'End With

For i = 1 To 2

    With wbBudget.Sheets(i)
    
    .Activate
    ActiveSheet.Unprotect ("RATSWVNXYRCMPIZ")
    
    lr = .Range("A" & Rows.count).End(xlUp).Row + 10
    
    For r = lr To 9 Step -1
    
        If .Cells(r, 1) = "" Then
        
        If .Cells(r, 2) <> "" And .Cells(r, 6) = "" Then .Cells(r - 1, 2) = .Cells(r - 1, 2) & .Cells(r, 2)
        
        .Cells(r, 1).EntireRow.Delete (xlShiftUp)
        
        End If
        
    Next
    
    End With
    
Next

End Sub

Sub CollectBudget()

Dim BudgetArr(1 To 50, 1 To 2)
Dim myarr(4) As Variant

'----------取得預算書主要項目-------

With wbBudget.Sheets(1)

r = 10
i = 1

Do Until .Cells(r, 1) = ""

    BudgetArr(i, 1) = .Cells(r, 2)

    i = i + 1
    r = r + 1

Loop

End With


'---------比較預算書相對應欄位--------


With wbBudget.Sheets(2)

    For i = 1 To 50
    
        If BudgetArr(i, 1) = "" Then
        
            BudgetArr(i, 2) = BudgetArr(i - 1, 2) + 1
        
            Exit For
            
        End If
    
        r = 9
        
        Do Until .Cells(r, 1) = ""
        
            Debug.Print .Cells(r, 1) & "," & .Cells(r, 2)
        
            If .Cells(r, 2) = BudgetArr(i, 1) Then
            
                BudgetArr(i, 2) = r
            
            End If
        
            r = r + 1
        
        Loop
    
    Next

End With

'--------從預算書取存放資料-----------

With shtBudget

    '.Range("A" & r_Main).Resize(2000, 30).Delete (xlShiftUp) '初始化資料

End With

continue:

For i = 1 To 50

    If BudgetArr(i, 1) = "" Then Exit For
    
    For r = BudgetArr(i, 2) To BudgetArr(i + 1, 2) - 1

    If r = BudgetArr(i, 2) Then
    
        Title = BudgetArr(i, 1)

    End If

        With wbBudget.Sheets(2)
        
            For j = 0 To 4
            
                myarr(j) = .Cells(r, j + 1)
            
            Next
            
        End With
        
        With shtBudget
        
            If Title = myarr(1) And myarr(4) = "" Then GoTo jump
            
            For j = 0 To UBound(myarr)
            
                .Cells(r_Main, j + 1) = myarr(j)
            
            Next
            
            '--------日報基本資料格式設定開始---------
            
            .Cells(r_Main, 6) = "=D" & r_Main & "*E" & r_Main
            
            .Range("A" & r_Main & ":F" & r_Main).HorizontalAlignment = xlLeft
            .Range("A" & r_Main & ":F" & r_Main).Borders.LineStyle = xlContinuous
            .Range("D" & r_Main & ":F" & r_Main).NumberFormatLocal = "0.00"
            .Range("D" & r_Main & ":E" & r_Main).Interior.ColorIndex = 40 '鎖定儲存格用
            .Range("D" & r_Main & ":F" & r_Main).Font.Name = "Times New Roman"
            
            '--------日報基本資料格式設定結束---------
            
            .Cells(r_Main, 1) = Title
            
            r_Main = r_Main + 1
 
        End With
        
jump:
        
    Next
    
Next

wbBudget.Close Savechanges:=False 'Ending結束預算編輯

End Sub

Sub ArrangeTitle() '設定準則

With shtBudget

    For r = 3 To .Cells(1, 1).End(xlDown).Row

        Title = .Cells(r, 1)
        
        Select Case Title
        
        Case "主體工程工作費": .Cells(r, 1).Font.ColorIndex = color_show: .Cells(r, 7) = color_show
        Case "主體工程": .Cells(r, 1).Font.ColorIndex = color_show: .Cells(r, 7) = color_show
        Case "雜項工程工作費": .Cells(r, 1).Font.ColorIndex = color_show: .Cells(r, 7) = color_show
        Case "雜項工程": .Cells(r, 1).Font.ColorIndex = color_show: .Cells(r, 7) = color_show
        Case "環境保護措施費": .Cells(r, 1).Font.ColorIndex = color_percent: .Cells(r, 7) = color_percent
        Case "職業安全衛生費": .Cells(r, 1).Font.ColorIndex = color_percent: .Cells(r, 7) = color_percent
        Case "廠商品質管制作業費": .Cells(r, 1).Font.ColorIndex = color_percent: .Cells(r, 7) = color_percent
        Case "包商管理費": .Cells(r, 1).Font.ColorIndex = color_percent: .Cells(r, 7) = color_percent
        Case "工程保險費": .Cells(r, 1).Font.ColorIndex = color_show: .Cells(r, 7) = color_show
        Case "營業稅(5%)": .Cells(r, 1).Font.ColorIndex = color_percent: .Cells(r, 7) = color_percent
        Case "空氣污染防制費": .Cells(r, 1).Font.ColorIndex = color_hide: .Cells(r, 7) = color_hide
        Case "工程管理費": .Cells(r, 1).Font.ColorIndex = color_hide: .Cells(r, 7) = color_hide
        Case "技師簽證費": .Cells(r, 1).Font.ColorIndex = color_hide: .Cells(r, 7) = color_hide
        Case "工程預備費": .Cells(r, 1).Font.ColorIndex = color_hide: .Cells(r, 7) = color_hide
        
        End Select

    Next

End With

End Sub

Sub ReArrangeTitle() '重整準則

With shtBudget

For r = 3 To .Cells(1, 1).End(xlDown).Row

    .Cells(r, 1).Font.ColorIndex = .Cells(r, 7)

Next

End With

End Sub

Sub CollectTitle() '收集不同的Title

With shtBudget

    collTitle.Add .Cells(3, 1).Value
    collTitleRow.Add 3
    collTitleRule.Add 3

    For r = 3 To .Cells(1, 1).End(xlDown).Row
    
        Title = .Cells(r, 1)
        TitleRule = .Cells(r, 7)
        
        IsExisted = False
    
        For j = 1 To collTitle.count
            
            ExistTitle = collTitle.Item(j)
            
            If Title = ExistTitle Then IsExisted = True
            
        Next
        
        If IsExisted = False Then
            collTitle.Add Title
            collTitleRow.Add r
            collTitleRule.Add TitleRule
        End If
    
    Next

End With

End Sub

Sub ClearOldReport() '請除舊有Main工作表報表資料

With shtMain

    lr = .Cells(Rows.count, c_Main).End(xlUp).Row
    
    .Cells(r_Main, c_Main).Resize(lr, 5).ClearContents

End With

End Sub

Sub ExportToReport() '匯出至Main

c_Main = 6
c_Diff = 6 - 2

With shtBudget

For i = 1 To collTitle.count

    Title = collTitle.Item(i)
    TitleRule = collTitleRule.Item(i)
    TitleRow = collTitleRow.Item(i)
    
    If i = collTitle.count Then
        NextTitleRow = TitleRow
    Else
        NextTitleRow = collTitleRow.Item(i + 1) - 1
    End If
    
    Select Case TitleRule
    
    Case color_show
    
        For r = TitleRow To NextTitleRow
        
            For c = c_Main To c_Main + 3
        
                shtMain.Cells(r_Main, c) = .Cells(r, c - c_Diff)
            
            Next
            
            shtMain.Cells(r_Main, c_Main + 4) = "=" & Chr(64 + c_Main + 2) & r_Main & "*" & Chr(64 + c_Main + 3) & r_Main
            
            shtMain.Cells(r_Main, c_Main).Font.ColorIndex = TitleRule
            
            r_Main = r_Main + 1
            
        Next
    
    Case color_percent
    
        sumMoney = 0
        j = 0
    
        For r = TitleRow To NextTitleRow
        
            sumMoney = sumMoney + .Cells(r, 6)
            
        Next
        
        arr = Array(.Cells(TitleRow, 1), "式", 1, sumMoney, .Cells(TitleRow, 4) * sumMoney)
        
        For c = c_Main To c_Main + 3
        
            shtMain.Cells(r_Main, c) = arr(j)
            
            j = j + 1
            
        Next
       
        shtMain.Cells(r_Main, c_Main + 4) = "=" & Chr(64 + c_Main + 2) & r_Main & "*" & Chr(64 + c_Main + 3) & r_Main
       
        shtMain.Cells(r_Main, c_Main).Font.ColorIndex = TitleRule
        
        r_Main = r_Main + 1
        
    End Select

Next

End With

End Sub
