VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Test 
   Caption         =   "試體設定"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9072.001
   OleObjectBlob   =   "frm_Test.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "frm_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()

Dim PCCES_obj As New clsPCCES
Dim Inf_obj As New clsInformation

rec_date = Inf_obj.startDate

Set coll = PCCES_obj.getAllItemsByRecDate(rec_date)

For Each it In coll

    Me.cboMatName.AddItem it
    Me.cboTestName.AddItem it

Next

End Sub
