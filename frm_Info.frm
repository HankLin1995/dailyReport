VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Info 
   Caption         =   "監造日報表VBA"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   11148
   OleObjectBlob   =   "frm_Info.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "frm_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CommandButton1_Click()
ERRORForm.Show
Unload Me
End Sub



Private Sub Image2_Click()
ActiveWorkbook.FollowHyperlink Address:="https://hankvba.blogspot.com", NewWindow:=True
End Sub

Private Sub Label12_Click()
ActiveWorkbook.FollowHyperlink Address:="https://hankvba.blogspot.com/2018/12/excel-vba-2.html", NewWindow:=True
End Sub

Private Sub Label14_Click()
ActiveWorkbook.FollowHyperlink Address:="https://docs.google.com/presentation/d/1JCvb9UrkNpCMiK6UO5loroEKePje5LWq/edit?usp=sharing&ouid=112944893851556117594&rtpof=true&sd=true", NewWindow:=True
End Sub

Private Sub Label15_Click()
ActiveWorkbook.FollowHyperlink Address:="https://creativecommons.org/licenses/by-nc/3.0/tw/legalcode", NewWindow:=True
End Sub
