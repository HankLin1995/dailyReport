VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_signup 
   Caption         =   "���U�u��"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4710
   OleObjectBlob   =   "frm_signup.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "frm_signup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub CommandButton1_Click()

If Not tboID.Text Like "*@*" Then

    MsgBox "Mail�榡�ܩ_��!", vbCritical
    Exit Sub
    
End If

If Len(tboPASSWORD.Text) < 8 Then

    MsgBox "�K�X���צܤ֤K��H�W!", vbCritical
    Exit Sub

End If

If Len(tboWG.Text) = 0 Then

    MsgBox "�Ч����g�����W��!", vbCritical
    Exit Sub
    
End If

If Len(tboName.Text) = 0 Then

    MsgBox "�ж�g���U�H���m�W!", vbCritical
    Exit Sub
    
End If

Dim obj As New clsFetchURL
Call obj.signup(tboID.Text, tboPASSWORD.Text, tboWG.Text, tboName.Text)

Unload frm_signup

End Sub

Private Sub UserForm_Click()

End Sub
