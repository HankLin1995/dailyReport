VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Photo_TMP 
   Caption         =   "UserForm1"
   ClientHeight    =   9060.001
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   6960
   OleObjectBlob   =   "frm_Photo_TMP.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "frm_Photo_TMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Image1_Click()

End Sub

Private Sub TextBox1_Change()

Me.Image1.Picture = LoadPicture(Me.TextBox1.Text)
Me.Image1.PictureSizeMode = 0 ' fmPictureSizeModeStretch

End Sub

Private Sub UserForm_Initialize()

End Sub
