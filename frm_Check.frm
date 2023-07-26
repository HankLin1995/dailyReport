VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Check 
   Caption         =   "��d�������"
   ClientHeight    =   9285.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11970
   OleObjectBlob   =   "frm_Check.frx":0000
   StartUpPosition =   1  '���ݵ�������
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

check_cnt = check_obj.countChecks(check_item)
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

arr = Array(check_item, check_item_eng, check_cnt, check_date, check_style, check_loc, , , photo_prompt)

Call myFunc.AppendData("Check", arr)

End With

Unload Me

End Sub

Private Sub CommandButton1_Click()

'f = "G:\�ڪ����ݵw��\ExcelVBA\�ʳy�����\�W�ү���\0109\113_200116_0026.jpg"

ChDir getThisWorkbookPath & "\�I�u�Ӥ�\"

If f = "" Then f = Application.GetOpenFilename

If f = "False" Then MsgBox "�����o�ɮ�", vbCritical: Exit Sub

Me.txtFilePath = f

End Sub

Private Sub CommandButton3_Click()

With lstCheckTable

If .ListIndex <> -1 Then Call .RemoveItem(.ListIndex)

End With

End Sub

Private Sub CommandButton2_Click()

If Me.txtFilePath Like "*>*" Then

    MsgBox "���|�����o�t��>", vbCritical
    Me.txtFilePath = ""
    Exit Sub

End If

If Me.txtPhotoNote Like "*>*" Then

    MsgBox "�����椤���o�t��>", vbCritical
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

If check_loc <> "" Then Me.txtCheckLoc = check_loc

End Sub

Private Sub lstCheckTable_Click()

With lstCheckTable

If .ListIndex <> -1 Then

    Me.txtFilePath = .List(.ListIndex, 0)
    Me.txtPhotoNote = .List(.ListIndex, 1)

End If

End With

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

Me.cboCheckStyle.AddItem "���簱�d�I"
Me.cboCheckStyle.AddItem "�I�u��d�I"
Me.cboCheckStyle.ListIndex = 0

With lstCheckTable

    .ColumnCount = 2
    .ColumnWidths = "60,120"

End With


End Sub
