VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MixData 
   Caption         =   "組合工項"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8130
   OleObjectBlob   =   "MixData.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "MixData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub cboItem_Change()

If Me.cboItem = "" Then Exit Sub

Dim obj As New clsBasicData

Call obj.RetrunUnit_Mix

End Sub

Private Sub cmdMixAdd_Click()

With Me

    .txtMixTable = .txtMixTable & "," & .cboItem & "," & .txtAmount & "," & .lblUnit

End With

End Sub

Private Sub cmdMixClear_Click()

Me.txtMixTable = ""

End Sub

Private Sub cmdOutput_Click()

With Me

    MixName = .txtMixName
    MixDefine = .txtDefine
    MixDefineTotal = .txtDefineTotal
    MixUnit = .txtUnit
    MixChannel = .txtChannel
    tmp = split(mid(.txtMixTable, 2), ",")
    
End With

Dim obj As New clsMixData

Call obj.AppendData(MixName, MixDefine, MixDefineTotal, MixUnit, MixChannel, tmp)

'With Sheets("Mix")
'
'    .UsedRange.EntireRow.Hidden = False '摺疊會出問題!
'
'    lr = .Cells(Rows.count, 4).End(xlUp).Row + 1
'
'    .Cells(lr, 1) = MixName
'    .Cells(lr, 2) = MixDefine
'    .Cells(lr, 3) = MixDefineTotal
'
'    For i = 0 To UBound(tmp) - 1 Step 3
'
'        .Cells(lr + j, 4) = tmp(0 + i)
'        .Cells(lr + j, 5) = tmp(1 + i)
'        .Cells(lr + j, 6) = tmp(2 + i)
'
'        j = j + 1
'
'    Next
'
'End With

End Sub


Private Sub txtMixName_AfterUpdate()

Dim Mix_obj As New clsMixData

If Mix_obj.IsMixNameUsed(Me.txtMixName) = True Then MsgBox "工項概要已經被使用!", vbCritical: Me.txtMixName = ""

End Sub

Private Sub txtMixName_Change()

End Sub

Private Sub UserForm_Initialize()

Dim obj As New clsBasicData

'obj.ReadData
obj.Init_Mix

End Sub
