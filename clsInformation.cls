VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private conName As String
Private ContractNum As String
Private ContractMoney As Double
Private Contractor As String
Private StartDate As Date
Private WorkDay As Integer
Private WorkDayAdd As Integer
Private EndDate As Date
Private Superviser As String
Private Man As String
Private Sta As String
Private IsChanged As Boolean
Private IsEnlarged As Boolean
Private EnlargedDate As Date

Dim collContractDay As New Collection
Dim collContractMoney As New Collection

Private shtMain As Object
Private shtDiary As Object

Private ConRow As Integer
Private TestRow As Integer
Private ChannelRow As Integer
Private ItemRow As Integer
Private PgsRow As Integer

Private Sub Class_Initialize()

Set shtMain = Sheets("Main")
Set shtDiary = Sheets("Diary")

Call GetBasicInformation '讀取Main工作表基本資訊

End Sub

Function ReturnInfoArr() '回傳Main工作表基本資訊

Dim arr As Variant

arrdata = Array(conName, ContractName, Contractor, StartDate, WorkDay, Superviser, ContractMoney, _
                      EnlargedDate, EndDate, IsEnlarged, IsChanged, collContractMoney, collContractDay, WorkDayAdd)
arrtitle = Array("ConName", "ContractName", "Contractor", "StartDate", "WorkDay", "Superviser", "ContractMoney", _
                      "EnlargedDate", "EndDate", "IsEnlarged", "IsChanged", "collContractMoney", "collContractDay", "WorkDayAdd")

ReDim arr(1 To UBound(arrdata) + 1, 1 To 2)

For i = 0 To UBound(arrdata)
    
    arr(i + 1, 1) = arrtitle(i)
    
    If arrtitle(i) Like "coll*" Then
        Set arr(i + 1, 2) = arrdata(i)
    Else
        arr(i + 1, 2) = arrdata(i)
    End If
    
Next

InfoArr = arr

ReturnInfoArr = InfoArr

End Function

Private Sub GetBasicInformation()

IsEnlarged = False
IsChanged = False

With shtMain

    conName = .Cells(1, 2)
    ContractName = .Cells(2, 2)
    Contractor = .Cells(3, 2)
    StartDate = .Cells(4, 2)
    WorkDay = .Cells(5, 2)
    Superviser = .Cells(7, 2)
    EnlargedDate = .Cells(6, 3)
    
    If EnlargedDate <> Empty Then
        IsEnlarged = True
        WorkDayAdd = .Cells(6, 2) - WorkDay
        WorkDay = .Cells(6, 2)
    End If

    EndDate = StartDate + WorkDay - 1

    For r = 10 To 12
    
        If .Cells(r, 3) <> "" Then
            
            collContractMoney.Add .Cells(r, 2).Value
            collContractDay.Add .Cells(r, 3).Value
            
            ContractMoney = .Cells(r, 2)
        Else
        
            Exit For
            
        End If
    
    Next
    
    If collContractDay.count > 1 Then IsChanged = True

End With

End Sub

Function GetStartDate()
    GetStartDate = StartDate
End Function

Function GetEndDate()
    GetEndDate = EndDate
End Function

Function GetcollContractDay()
    Set GetcollContractDay = collContractDay
End Function

Sub ProgressNew() '計算各天數進度傳入Diary工作表

Dim collDay As New Collection
Dim collPer As New Collection
Dim arr() As Variant

Call GetMainRow

'----------錯誤檢查----------

IsCheckPass = ProgressCheck

If IsCheckPass = False Then Exit Sub

'----------計算資料----------

With shtMain

lr = .Cells(PgsRow, 1).End(xlDown).Row

ReDim arr(1 To WorkDay, 1 To 2)

For r = PgsRow + 1 To lr

    pday = .Range("A" & r)
    p = .Range("B" & r)
    
    If r = lr Then
    
        arr(j, 1) = pday
        arr(j, 2) = Round(p, 3)
    
        GoTo X
    
    End If
    
    DuringDay = .Range("A" & r + 1) - pday
    
    Add = (.Range("B" & r + 1) - p) / DuringDay
    
    jj = 1

    For j = 1 + jc To DuringDay + jc
    
        arr(j, 1) = pday + (jj - 1)
        arr(j, 2) = Round(p + Add * (jj - 1), 3)
        
        jj = jj + 1

        jc = jc + 1
    
    Next

Next

End With

'----------匯入資料----------

X:

With shtDiary


    For u = 1 To jc + 1

        r = u + 1

   'r = 2
    'u = 1
    
    .Cells(r, 1) = u
    .Cells(r, 2) = arr(u, 1)
    .Cells(r, 3) = "晴"
    .Cells(r, 4) = arr(u, 2)
    .Cells(r, 4).NumberFormatLocal = "0.00%"
    
    Next
    
'    Do Until .Cells(r, 2) = ""
'
'    .Cells(r, 4) = arr(u, 2)
'
'    .Cells(r, 4).NumberFormatLocal = "0.00%"
'
'    Loop

End With

MsgBox "進度取得完畢，請至""Diary""查看!!"

End Sub

Private Function ProgressCheck() '確認進度表有無異常項目

ProgressCheck = True

With shtMain

    lr = .Range("A" & PgsRow).End(xlDown).Row

    comfinalday = CDate(.Cells(lr, 1))
    
    If EndDate <> comfinalday Then
    
        ProgressCheck = False
        
        MsgBox "正確結束日期應該為 " & EndDate
        
        Exit Function

    End If
    
    For r = PgsRow + 1 To lr - 1
    
        beforeDate = CDate(.Cells(r, 1))
        afterDate = CDate(.Cells(r + 1, 1))
    
        If afterDate < beforeDate Then
        
                ProgressCheck = False
                MsgBox "你的日期前後順序有問題!!"
                'Exit Function
        
        End If
    
        If .Cells(r + 1, 2) < .Cells(r, 2) Then
        
                ProgressCheck = False
                MsgBox "你的累積進度有問題!!"
                Exit Function

        End If
        
        If r = lr - 1 And .Cells(r + 1, 2) <> 1 Then
        
            ProgressCheck = False
            
            MsgBox "最後總和不等於100%,請再檢查看看"
            
            Exit Function
            
        End If
    
    Next

End With

End Function

Private Sub GetMainRow() '取得各資料小分類首列Row

Call ReturnMainRow(arr)

ConRow = arr(0)
TestRow = arr(1)
ChannelRow = arr(2)
ItemRow = arr(3)
PgsRow = arr(4)

End Sub
