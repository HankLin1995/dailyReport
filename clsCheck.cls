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

'Set shtCheckTMP = Workbooks.Open(getThisWorkbookPath & "\查驗TMP\查驗.xls")

'targetCheck.Add "開挖", "開挖"
'targetCheck.Add "土方", "土方"
'targetCheck.Add "鋼筋", "鋼筋"
'targetCheck.Add "模板", "模板"
'targetCheck.Add "混凝土", "混凝土"
'targetCheck.Add "混凝土完成", "混凝土完成"
'targetCheck.Add "假設工程", "假設工程"
'targetCheck.Add "碎石級配", "碎石級配"
'targetCheck.Add "瀝青混凝土", "瀝青混凝土"
'targetCheck.Add "傾倒式閘門", "傾倒式閘門"
'targetCheck.Add "吊門機製造及安裝", "吊門機製造及安裝"
'targetCheck.Add "自動測報", "自動測報"
'targetCheck.Add "電氣", "電氣"
'targetCheck.Add "焊接", "焊接"

End Sub

Sub CheckList()

Set fso = CreateObject("Scripting.FileSystemObject")

With Sheets("Check")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 3 To lr

    checkName_eng = .Cells(r, 2)
    checkNum = .Cells(r, 3)
    
    wbpath = getThisWorkbookPath & "\抽查表Output\" & checkName_eng & "-" & checkNum & ".xls"
    
    If fso.fileexists(wbpath) = True Then
        .Cells(r, "G") = "Printed"
    Else
        .Cells(r, "G") = ""
    End If

Next

End With
End Sub

Sub printCheckTable() '列印查驗表

'Application.ScreenUpdating = False

Dim print_obj As New clsPrintOut
Dim Inf_obj As New clsInformation

conName = Inf_obj.conName

With ThisWorkbook.Sheets("Check")

    For r = 3 To .Cells(.Rows.Count, 1).End(xlUp).Row
    
        '==============基本資料取入===============
    
        'shtName = .Cells(r, 1)
        checkName = .Cells(r, 1)
        checkName_eng = .Cells(r, 2)
        checkNum = .Cells(r, 3)
        CheckDate = .Cells(r, 4)
        checkloc = .Cells(r, "F")
        CheckType = .Cells(r, "E")
        CheckLog = .Cells(r, "G")
        CheckPrintAgain = .Cells(r, "H")
        
        If CheckLog <> "" And CheckPrintAgain <> "Y" Then GoTo PrintPass
        
        .Cells(r, "G") = "Printed"
        
        If CheckType = "檢驗停留點" Then
            prompt = "■檢驗停留點                       □施工抽查點"
        Else
            prompt = "□檢驗停留點                       ■施工抽查點"
        End If

        '==============從查驗TMP裡面取範本copy到成果=======
        
        fs_name = checkName & "[" & checkName_eng & "]"
        fs = getThisWorkbookPath & "\抽查表TMP\" & fs_name & ".xls"
        
        Set wb = Workbooks.Open(fs) '範例檔案
        
        Set sht = wb.Sheets(1)
        
        a = CollectRange(sht)
        
        With wb.Sheets(1)
        
            If a(0) <> "" Then .Range(a(0)) = checkNum
            If a(1) <> "" Then .Range(a(1)) = conName
            If a(2) <> "" Then .Range(a(2)) = checkloc
            If a(3) <> "" Then .Range(a(3)) = CheckDate
            If a(4) <> "" Then .Range(a(4)) = prompt
        
        End With
        
        Call print_obj.SpecificShtObjToXLS(sht, checkName_eng & "-" & checkNum, getThisWorkbookPath & "\抽查表Output\" & checkName_eng & "-" & checkNum)
        
        p = p & checkName & ">" & checkName_eng & "-" & checkNum & vbNewLine
         
        wb.Close False
        
PrintPass:

    Next

    
End With

If p <> "" Then
    MsgBox "抽查記錄表列印明細:" & vbNewLine & p, vbInformation
Else
    MsgBox "抽查記錄表皆已列印完畢!", vbInformation
End If
'Application.ScreenUpdating = True

End Sub

Sub collectFilesBySelect()

Application.ScreenUpdating = False

Set coll_path_sort = getFilePathBySelect

Set wb_combine = Workbooks.Add

For Each filePath In coll_path_sort

    filename = mid(filePath, InStrRev(filePath, "\") + 1)
    fileExtension = mid(filename, InStrRev(filename, ".") + 1)
    
    filename_short = mid(filename, 1, InStrRev(filename, ".") - 1)
    
    Set wb_check = Workbooks.Open(filePath)
    Set sht = wb_check.Sheets(1)
    sht.Name = filename_short
    
    sht.Copy before:=wb_combine.Sheets(wb_combine.Sheets.Count)
    
    wb_check.Close False
    
    Set wb_check = Nothing
Next

Application.DisplayAlerts = False
wb_combine.Sheets("工作表1").Delete
wb_combine.Activate
Application.DisplayAlerts = True

Application.ScreenUpdating = True

End Sub

Private Function getFilePathBySelect()

Dim coll_rows As New Collection
Dim coll_path_sort As New Collection

For Each rng In Selection

On Error Resume Next
coll_rows.Add rng.Row, CStr(rng.Row)
On Error GoTo 0

Next

For Each r In coll_rows

    myPath = getThisWorkbookPath & "\抽查表Output\" & Sheets("Check").Cells(r, 2) & "-" & Sheets("Check").Cells(r, 3) & ".xls"

    coll_path_sort.Add myPath

Next

Set getFilePathBySelect = coll_path_sort

End Function


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

'=======202307=====


Function getCheckFileNames()

Dim coll As New Collection

myFolder = getThisWorkbookPath & "\抽查表TMP" ' getSavedFolder

'MsgBox myFolder

Set coll_path = GetFilePathsInFolder(myFolder)

For Each filePath In coll_path

    filename = mid(filePath, InStrRev(filePath, "\") + 1)
    pt = InStrRev(filename, ".")
    filename_short = mid(filename, 1, pt - 1)

    coll.Add filename_short

Next

Set getCheckFileNames = coll

End Function


Function countChecks(ByVal check_name As String)

Dim myFunc As New clsMyfunction

Set coll_rows = myFunc.getRowsByUser2("Check", check_name, 2, "抽查表名稱")

countChecks = coll_rows.Count

End Function
