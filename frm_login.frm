VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_login 
   Caption         =   "Login"
   ClientHeight    =   5310
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4875
   OleObjectBlob   =   "frm_login.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub CommandButton1_Click() '�n�J

Dim obj As New clsFetchURL

id = Me.tboID
password = Me.tboPASSWORD

Call obj.checkAccesByID(id, password)

Unload Me

End Sub

Private Sub CommandButton2_Click() '���U

frm_signup.Show

End Sub

Private Sub CommandButton3_Click() '�ѰO�K�X

Dim obj As New clsFetchURL

If Me.tboID = "" Then MsgBox "�b�����o���ŭ�!!!", vbCritical: Exit Sub

obj.getPassword (Me.tboID)

End Sub

Private Sub UserForm_Terminate()

If Me.Label6 <> "Pass" Then ThisWorkbook.Close SaveChanges:=False

End Sub
