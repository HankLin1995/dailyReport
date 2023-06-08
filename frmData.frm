VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmData 
   Caption         =   "填寫施作資料"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10170
   OleObjectBlob   =   "frmData.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboItem_Change()

Call ResetItem

End Sub

Private Sub CheckBox1_Click() '20221125 僅提供剩餘數量

Me.cboItem.Clear

Set coll = getCollNotZero

For Each it In coll

    Me.cboItem.AddItem (it)

Next

End Sub

Private Sub chkUntilDay_Click()

Call ResetItem

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

Private Sub cmdDetail_Click()

On Error GoTo ERRORHANDLE:

Dim obj As New clsDetail

With frm_Detail

    Item = Me.cboItem.Text
    
    Call obj.getPropertiesByName(Item)
    Call obj.setItemToCbo

    .Label15.Caption = Item
    .Show

End With

Exit Sub

ERRORHANDLE:
MsgBox "Detail找不到" & Item & "!", vbCritical

End Sub

Private Sub cmdOutput_Click()

Dim obj2 As New clsDetail
obj2.getPropertiesByName (frmData.cboItem.Text)
obj2.setAmount (frmData.txtDetailTable.Text)

Dim obj As New clsRecord

obj.ReadData
obj.Recording

If frmData.txtDetailTable <> "" Then obj.Recording_Detail

frmData.cboItem = ""

Call ResetItem

End Sub


Private Sub txtAmount_Change()

With frmData

    amount = .txtAmount
    
    If IsNumeric(amount) And amount > 0 Then
        .lblLast = .lblStore - .txtAmount
    Else
        .lblLast = .lblStore
    End If

End With

End Sub

Private Sub txtDay_Change()

Call ResetItem

End Sub

Private Sub UserForm_Initialize()

Dim obj As New clsBasicData

obj.ReadData
obj.Init

Dim obj2 As New clsCheck

Call obj2.AddCheckTable(frmData)

End Sub
