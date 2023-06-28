VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MixData_Main 
   Caption         =   "填寫組合工項"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9915.001
   OleObjectBlob   =   "MixData_Main.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "MixData_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboItem_Change()

Call ResetItem_Mix

End Sub

Private Sub chkUntilDay_Click()

Call ResetItem_Mix

End Sub


Private Sub cmdCheckAdd_Click()

ImpCheck = "施工抽查點"

If MsgBox("是否為查驗停留點?", vbYesNo) = vbYes Then ImpCheck = "查驗停留點"

With Me

    .txtCheckTable = .txtCheckTable & "," & .cboCheck & "," & ImpCheck
    
    .cboCheck = ""

End With

End Sub

Private Sub cmdCheckClear_Click()

    Me.txtCheckTable.Text = ""

End Sub

Private Sub cmdOutput_Click()

Dim obj As New clsRecord

obj.ReadData_Mix
obj.Recording_Mix

Call ResetItem_Mix

End Sub

Private Sub CommandButton2_Click() '20230205

strLoc = txtWhere.Text

If strLoc Like "*、*" Then

    loc_tmp = Split(strLoc, "、")
    
    For Each it In loc_tmp
    
        sumL = sumL + calcLoc(it)
        
    Next

Else

    sumL = calcLoc(strLoc)

End If

Debug.Print sumL

txtAmount = sumL

End Sub

Private Sub txtAmount_Change()

With MixData_Main

    amount = .txtAmount
    
    If IsNumeric(amount) And amount > 0 Then
        .lblLast = .lblStore - .txtAmount
    Else
        .lblLast = .lblStore
    End If

End With

End Sub

Private Sub txtDay_Change()

Call ResetItem_Mix

End Sub

Private Sub UserForm_Initialize()

Dim obj As New clsMixData

obj.ReadData
obj.ReadMainData
obj.Init

Dim obj2 As New clsCheck

Call obj2.AddCheckTable(MixData_Main)

Me.txtDay = Format(Now(), "yyyy/mm/dd")

End Sub
