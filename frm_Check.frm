VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Check 
   Caption         =   "抽查紀錄表單"
   ClientHeight    =   9285.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11970
   OleObjectBlob   =   "frm_Check.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "frm_Check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub cmdOutput_Click()

Dim myFunc As New clsMyfunction
Dim check_obj As New clsCheck

With Me

check_name = .cboCheckItem

Call myFunc.splitFileName_Check(check_name, check_item, check_item_eng)

check_cnt = check_obj.countChecks(check_item) + 1
check_date = .txtCheckDate
check_style = .cboCheckStyle
check_loc = .txtCheckLoc
check_photo_inf = .lstCheckTable.List

cnt = UBound(check_photo_inf, 1)

For i = 0 To cnt

    photo_path = check_photo_inf(i, 0)
    photo_note = check_photo_inf(i, 1)

    photo_prompt = photo_prompt & photo_path & ">" & photo_note & ","

Next

If cnt <> -1 Then photo_prompt = mid(photo_prompt, 1, Len(photo_prompt) - 1)

arr = Array(check_item, check_item_eng, check_cnt, check_date, check_style, check_loc, , , photo_prompt)

Stop

If .Label_Row = "" Then
    r = 0
Else
    r = CInt(.Label_Row)
End If

If r <> 0 Then

    check_cnt = Sheets("Check").Range("C" & r)
    arr = Array(check_item, check_item_eng, check_cnt, check_date, check_style, check_loc, , , photo_prompt)
    Sheets("Check").Range("A" & r & ":I" & r) = arr
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(getThisWorkbookPath & "\抽查表Output\" & check_item_eng & "-" & check_cnt & ".xls") Then
        Set f = fso.getFile(getThisWorkbookPath & "\抽查表Output\" & check_item_eng & "-" & check_cnt & ".xls")
        Kill f
    End If

Else

    Call myFunc.AppendData("Check", arr)

End If

End With

Unload Me

If r <> 0 Then Call cmdPrintCheck

End Sub



Private Sub CommandButton1_Click()

'f = "G:\我的雲端硬碟\ExcelVBA\監造日報表\上課素材\0109\113_200116_0026.jpg"

' 設定起始資料夾的路徑

Dim myFunc As New clsMyfunction
Dim initialFolder As String

If Me.txtFilePath = "" Then

initialFolder = getThisWorkbookPath

Else

initialFolder = mid(Me.txtFilePath, 1, InStrRev(Me.txtFilePath, "\"))

End If

If f = "" Then f = myFunc.FileOpen(initialFolder, "JPG(*.jpg)", "*.jpg") '.GetOpenFilename

If f = "False" Then MsgBox "未取得檔案", vbCritical: Exit Sub

Me.txtFilePath = f

End Sub

Private Sub CommandButton3_Click()

With lstCheckTable

If .ListIndex <> -1 Then Call .RemoveItem(.ListIndex)

End With

End Sub

Private Sub CommandButton2_Click()

If Me.txtFilePath Like "*>*" Then

    MsgBox "路徑中不得含有>", vbCritical
    Me.txtFilePath = ""
    Exit Sub

End If

If Me.txtPhotoNote Like "*>*" Then

    MsgBox "說明欄中不得含有>", vbCritical
    Me.txtPhotoNote = ""
    Exit Sub

End If

With lstCheckTable

i = UBound(.List)
.AddItem ""

.List(i + 1, 0) = Me.txtFilePath
.List(i + 1, 1) = Me.txtPhotoNote

End With

End Sub



Private Sub CommandButton4_Click()

Dim REC_obj As New clsRecord

check_date = CDate(Me.txtCheckDate)

check_loc = REC_obj.getExistLocByRecDate(check_date)

If check_loc Like "*,*" Then

    tmp = split(check_loc, ",")
    
    Me.txtCheckCanal = tmp(0)
    Me.txtCheckLocDetail = tmp(1)
    Me.txtCheckLoc = check_loc

End If

End Sub

Private Sub lstCheckTable_Click()

With lstCheckTable

If .ListIndex <> -1 Then

    Me.txtFilePath = .List(.ListIndex, 0)
    Me.txtPhotoNote = .List(.ListIndex, 1)

End If

End With

End Sub



Private Sub txtCheckCanal_AfterUpdate()

Me.txtCheckLoc = Me.txtCheckCanal & "," & Me.txtCheckLocDetail

End Sub

Private Sub txtCheckLocDetail_AfterUpdate()

Me.txtCheckLoc = Me.txtCheckCanal & "," & Me.txtCheckLocDetail

End Sub



Private Sub txtFilePath_Change()

Me.Image1.Picture = LoadPicture(Me.txtFilePath)
Me.Image1.PictureSizeMode = fmPictureSizeModeStretch

End Sub

Private Sub UserForm_Initialize()

Dim check_obj As New clsCheck

Set coll_check_files = check_obj.getCheckFileNames

For Each it In coll_check_files

    Me.cboCheckItem.AddItem it, j

    j = j + 1

Next

Me.txtCheckDate = CDate(Format(Now(), "yyyy/mm/dd"))

Me.cboCheckStyle.AddItem "檢驗停留點"
Me.cboCheckStyle.AddItem "施工抽查點"
Me.cboCheckStyle.ListIndex = 0

With lstCheckTable

    .ColumnCount = 2
    .ColumnWidths = "60,120"

End With


End Sub
