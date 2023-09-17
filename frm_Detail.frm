VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Detail 
   Caption         =   "一式計算項目"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7545
   OleObjectBlob   =   "frm_Detail.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "frm_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub cboItem_Change()

subItem = Me.cboItem.Text

Dim obj As New clsDetail
Call obj.getPropertiesByName(Me.Label15.Caption)

Me.lblUnit = obj.getUnit(subItem)
Me.lblLast = obj.getLast(subItem)
Me.lblStore = Me.lblLast

End Sub

Private Sub cmdOutput_Click()

frmData.txtAmount = Me.txtRatio
frmData.txtDetailTable = Me.txtDetailTable

Unload Me

End Sub

Private Sub CommandButton1_Click()

Me.txtDetailTable = ""

End Sub

Private Sub CommandButton2_Click()

Dim obj As New clsDetail

mainItem = Me.Label15.Caption
subItem = Me.cboItem.Text

'========IsExisted==============

If Me.txtDetailTable Like "*" & subItem & "*" Then MsgBox "已經填寫過了!", vbCritical: Exit Sub

'========initialize=============

Call obj.getPropertiesByName(mainItem)

'=========if minus=========

If CInt(Me.lblLast) < 0 Then

    MsgBox "數量不得為負值!", vbCritical
    
    Me.lblLast = obj.getLast(subItem)
    Me.txtAmount = 0
    
    Exit Sub

End If

'===========ok=============

With Me

    .txtDetailTable = .txtDetailTable & "," & .cboItem & "," & .txtAmount & "," & .lblUnit
    .txtRatio = Round(obj.calcRatio(.txtDetailTable.Text), 2)
    
End With

End Sub

Private Sub txtAmount_Change()

With Me

    amount = .txtAmount

    If IsNumeric(amount) And amount > 0 Then
        .lblLast = .lblStore - .txtAmount
    Else
        .lblLast = .lblStore
    End If

End With

End Sub
