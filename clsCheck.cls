VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private shtCheck As Object
Private shtRec As Object
Private targetCheck As New Collection

Private Sub Class_Initialize()

Set shtCheck = Sheets("Check")
Set shtRec = Sheets("Records")

'Set shtCheckTMP = Workbooks.Open(ThisWorkbook.Path & "\查驗TMP\查驗.xls")

targetCheck.Add "開挖", "開挖"
targetCheck.Add "土方", "土方"
targetCheck.Add "鋼筋", "鋼筋"
targetCheck.Add "模板", "模板"
targetCheck.Add "混凝土", "混凝土"
targetCheck.Add "混凝土完成", "混凝土完成"
targetCheck.Add "假設工程", "假設工程"
targetCheck.Add "碎石級配", "碎石級配"
targetCheck.Add "瀝青混凝土", "瀝青混凝土"
targetCheck.Add "傾倒式閘門", "傾倒式閘門"
targetCheck.Add "吊門機製造及安裝", "吊門機製造及安裝"
targetCheck.Add "自動測報", "自動測報"
targetCheck.Add "電氣", "電氣"
targetCheck.Add "焊接", "焊接"

End Sub

Sub CheckList()

Dim collchn As New Collection

wbpath = ThisWorkbook.Path & "\查驗表Output\"

f = Dir(wbpath)

Do Until f = ""
    
    ptdot = InStr(1, f, ".")
    
    chn = mid(f, 1, ptdot - 1)
    
    collchn.Add chn, chn
    
    f = Dir()

Loop

Call CheckExist(collchn)

End Sub

Private Sub CheckExist(ByVal collchn As Object)

With Sheets("Check")

lr = .Cells(Rows.count, 1).End(xlUp).Row

For r = 2 To lr

    testchn = .Cells(r, 1) & "-" & .Cells(r, 3)
    
    On Error Resume Next
    
    a = collchn(testchn)
    
    .Cells(r, "G") = "Printed"
    
    If Err <> 0 Then .Cells(r, "G") = ""

Next

End With

End Sub

Sub AddCheckTable(ByVal frm As UserForm)   'frmData查驗表單內容

With frm

    .cboCheck.AddItem "開挖(放樣)工程查驗表"
    .cboCheck.AddItem "土方工程施工查驗表"
    .cboCheck.AddItem "鋼筋施工組立作業查驗表"
    .cboCheck.AddItem "模板施工組立作業查驗表"
    .cboCheck.AddItem "混凝土施工作業查驗表"
    .cboCheck.AddItem "假設工程施工查驗表"
    .cboCheck.AddItem "碎石級配工程施工查驗表"
    .cboCheck.AddItem "瀝青混凝土工程施工查驗表"
    .cboCheck.AddItem "傾倒式閘門工程施工查驗表"
    .cboCheck.AddItem "吊門機製造及安裝工程施工查驗表"
    .cboCheck.AddItem "自動測報工程施工查驗表"
    .cboCheck.AddItem "電氣工程施工查驗表"
    .cboCheck.AddItem "焊接工程施工查驗表"

End With

End Sub

Sub ExportToCheck() 'Records工作表匯入至Check工作表

With shtCheck

    lr = .Cells(Rows.count, 1).End(xlUp).Row
    If lr > 2 Then .Cells(2, 1).Resize(lr, 7).ClearContents '強制清除

End With

With shtRec

    sr = 3
    lr = .Cells(Rows.count, 1).End(xlUp).Row

    For r = sr To lr
    
        checkday = .Cells(r, 2)
        CheckLoc = .Cells(r, 3) & "," & .Cells(r, 4)
        If .Cells(r, 4) = "" Then CheckLoc = .Cells(r, 3)
        'CheckedLog = .Cells(r, "J")
        CheckType_origin = .Cells(r, "H")
        CheckImp_origin = .Cells(r, "I")
        
        'If CheckedLog <> "" Then GoTo KEEPHANDLE
        
        CheckType_tmp = Split(CheckType_origin, ",")
        CheckImp_tmp = Split(CheckImp_origin, ",")
    
        'If UBound(CheckType_tmp) <> -1 And CheckedLog = "" Then .Cells(r, "J") = Format(Now, "yyyy / mm / dd")
        
        For i = LBound(CheckType_tmp) To UBound(CheckType_tmp)
        
            CheckType = CheckType_tmp(i)
            CheckImp = CheckImp_tmp(i)
            
            arr = Array(CheckType, checkday, CheckImp, CheckLoc)
            
            Call WriteToCheck(arr)
            
            If CheckType = "混凝土" Then
                
                arr(0) = "混凝土完成"
                Call WriteToCheck(arr)
                
            End If
            
        Next
        
KEEPHANDLE:
    
    Next
    
End With

End Sub

Private Sub WriteToCheck(ByVal arr As Variant) '將arr資料帶入至Check工作表

CheckType = arr(0)
checkday = arr(1)
CheckImp = arr(2)
CheckLoc = arr(3)

Select Case CheckImp

Case 0: CheckImp_Ch = "施工抽查點"

Case 1: CheckImp_Ch = "檢驗停留點"

End Select

With shtCheck '取消查驗表(代碼)填入

    arr = Array(CheckType, checkday, CheckImp, CheckLoc)

    lr_data = .Cells(Rows.count, 1).End(xlUp).Row + 1
    
    .Range("A" & lr_data) = CheckType
    .Range("D" & lr_data) = checkday
    .Range("E" & lr_data) = CheckImp_Ch
    .Range("F" & lr_data) = CheckLoc

End With

End Sub

Sub CountCheck() '計算各查驗表編號

With shtCheck

    sr = 2
    lr = .Cells(Rows.count, 1).End(xlUp).Row
    
    For i = 1 To targetCheck.count
    
        j = 1
        
        For r = sr To lr
        
           If .Cells(r, 1) = targetCheck(i) Then
           
                .Cells(r, 3) = j
                j = j + 1
                
           End If
           
        Next
        
    Next

End With

End Sub

'Sub PrintCheckTable() '列印查驗表
'
'msg1 = Val(InputBox("請選擇列印方式" & vbNewLine & "1.PDF" & vbNewLine & "2.實體紙張"))
'
'If msg1 = 1 Then
'
'ElseIf msg1 = 2 Then
'
'    msg = MsgBox("目前的印表機為:" & Application.ActivePrinter & vbNewLine & "是否要繼續列印?", vbYesNo + vbInformation)
'    If msg = vbNo Then End
'
'Else
'    End
'End If
'
'With shtCheck
'
'lr = .Cells(Rows.count, 1).End(xlUp).Row
'
'For r = 2 To lr
'
'    CheckName = .Cells(r, 1)
'    CheckNum = .Cells(r, 3)
'    CheckTime = .Cells(r, 4)
'    CheckImp = .Cells(r, 5)
'    CheckLoc = .Cells(r, 6)
'    CheckLog = .Cells(r, "G")
'    CheckPrintAgain = .Cells(r, "H")
'
'    If CheckLog <> "" And CheckPrintAgain = "" Then GoTo PassHandler
'
'    Set sht = Sheets(CheckName)
'
'    With sht
'
'        .Visible = True
'
'        arr = CollectRange(sht)
'
'        If CheckName = "混凝土完成" Then
'
'            arr(4) = ""
'            .Range("E6") = "(檢查處混凝土澆置日期：" & CheckTime & ")"
'
'        End If
'
'        .Range(arr(0)) = CheckNum
'        .Range(arr(1)) = Sheets("Main").Range("B1")
'        .Range(arr(2)) = CheckLoc
'
'        If arr(4) <> "" Then .Range(arr(3)) = CheckTime
'
'        Select Case CheckImp
'
'        Case "檢驗停留點": prompt = "■查驗停留點                       □施工抽查點"
'        Case "施工抽查點": prompt = "□查驗停留點                       ■施工抽查點"
'
'        End Select
'
'        If arr(4) <> "" Then .Range(arr(4)) = prompt
'
'        Dim obj As New clsPrintOut
'
'        Call obj.ToPDF_Check(sht, CheckName & "-" & CheckNum, msg1)
'
'        .Visible = False
'
'PassHandler:
'
'    End With
'
'Next
'
'.Activate
'
'End With
'
'End Sub

Sub PrintCheckTable() '列印查驗表

fs = ThisWorkbook.Path & "\查驗TMP\查驗.xls"

Set wb = Workbooks.Open(fs) '範例檔案
Set wbnew = Workbooks.Add '成果檔案

conName = ThisWorkbook.Sheets("Main").Range("B1")

With ThisWorkbook.Sheets("Check")

    lr = .Cells(.Rows.count, 1).End(xlUp).Row

    For r = 2 To lr
    
        '==============基本資料取入===============
    
        shtName = .Cells(r, 1)
        CheckNum = .Cells(r, 3)
        CheckDate = .Cells(r, 4)
        CheckLoc = .Cells(r, "F")
        CheckType = .Cells(r, "E")
        CheckLog = .Cells(r, "G")
        CheckPrintAgain = .Cells(r, "H")
        
        If CheckLog <> "" And CheckPrintAgain = "" Then GoTo PassHandler
        
        .Cells(r, "G") = "Printed"
        
'        prompt = "□施工前                □施工中檢查              □施工完成檢查 "
        
'        If CheckType = "檢驗停留點" Then
'            prompt = "■查驗停留點                       □施工抽查點"
'        Else
'            prompt = "□查驗停留點                       ■施工抽查點"
'        End If

        '==============從查驗TMP裡面取範本copy到成果=======

        Application.DisplayAlerts = False
        wb.Sheets(shtName).Copy wbnew.Sheets(1)
        Application.DisplayAlerts = True
        
        Set shtCheckNew = wbnew.Sheets(1)
        
        a = CollectRange(shtCheckNew)
        
        With shtCheckNew
        
            If a(0) <> "" Then .Range(a(0)) = CheckNum
            If a(1) <> "" Then .Range(a(1)) = conName
            If a(2) <> "" Then .Range(a(2)) = CheckLoc
            If a(3) <> "" Then .Range(a(3)) = CheckDate
            If a(4) <> "" Then .Range(a(4)) = prompt
        
            .Name = shtName & "-" & CheckNum
        
        End With
        
PassHandler:
        
    Next
    
    '==================刪除成果內的原始工作表====================
    
    Application.DisplayAlerts = False
    wbnew.Sheets(wbnew.Sheets.count).Delete
    Application.DisplayAlerts = True

    wb.Close False

End With

End Sub

Private Function CollectRange(ByVal sht As Worksheet) '取得查驗表填寫欄位

Dim arr(4) As String

With sht

    For Each rng In .UsedRange
    
        If rng.Value = "%%頁碼%%" Then
            arr(0) = FindKeyAddress(rng)
        ElseIf rng.Value = "%%工程名稱%%" Then
            arr(1) = FindKeyAddress(rng)
        ElseIf rng.Value = "%%地點%%" Then
            arr(2) = FindKeyAddress(rng)
        ElseIf rng.Value = "%%時間%%" Then
            arr(3) = FindKeyAddress(rng)
        ElseIf rng.Value = "%%查驗形式%%" Then
            arr(4) = FindKeyAddress(rng) ': Exit For
        End If
    
    Next

End With

CollectRange = arr

End Function

Private Function FindKeyAddress(ByVal rng As Range) '針對特定儲存格取其有效輸入位址

If rng.MergeCells = True Then

    Set ma = rng.MergeArea
    tmp = Split(ma.Address, ":")
    
    Row1 = Val(mid(tmp(0), InStr(2, tmp(0), "$") + 1))
    Row2 = Val(mid(tmp(1), InStr(2, tmp(1), "$") + 1))
    
    Set rng_last = Range(tmp(0)) '.Offset(Row1 - Row2, 1)

Else

    Set rng_last = rng
    
End If

FindKeyAddress = rng_last.Address

End Function


