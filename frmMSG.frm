VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMSG 
   Caption         =   "���@�u�{�I�u��x"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6390
   OleObjectBlob   =   "frmMSG.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton1_Click()

If Me.Label1.caption = "�ݨ��լd" Then

ERRORForm.Show
Unload Me

End If

End Sub

Private Sub CommandButton2_Click()

Unload Me 'ERRORForm

End Sub

Private Sub UserForm_Click()

End Sub
