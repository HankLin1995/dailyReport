VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Report 
   Caption         =   "�^�X���"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   OleObjectBlob   =   "frm_Report.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "frm_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub CommandButton1_Click()

Dim obj As New clsFetchURL
obj.getReport (TextBox1.Text)

Unload Me

End Sub

Private Sub CommandButton2_Click()

Unload Me

End Sub

Private Sub UserForm_Initialize()

s = split(Application.StatusBar, ",")

account = mid(s(1), 5)

Me.Label7 = account
Me.Label6 = "Hi!" & account & ",�A���^���O�ڳЧ@���̤j�ʤO!"


End Sub
