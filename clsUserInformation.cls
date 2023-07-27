VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUserInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Function getMacAddress()

Dim objVMI As Object
Dim vAdptr As Variant
Dim objAdptr As Object
'Dim adptrCnt As Long


Set objVMI = GetObject("winmgmts:\\" & "." & "\root\cimv2")
Set vAdptr = objVMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

For Each objAdptr In vAdptr
    If Not IsNull(objAdptr.MACAddress) And IsArray(objAdptr.IPAddress) Then
        For adptrCnt = 0 To UBound(objAdptr.IPAddress)
        If Not objAdptr.IPAddress(adptrCnt) = "0.0.0.0" Then
            GetNetworkConnectionMACAddress = objAdptr.MACAddress
            Exit For
        End If
        Next
    End If
Next

getMacAddress = GetNetworkConnectionMACAddress

End Function

Sub Login(ByVal mac_add As String)

'Call showCmd

With Sheets("授權")
    
    'mac_add = getMacAddress
    'Debug.Print mac_add
    
    'If mac_add = "" Then MsgBox "未偵測到有網路!", vbCritical: Call hideCmd: Exit Sub 'ThisWorkbook.Close Savechanges:=False
    
    If checkIsExist(mac_add) = False Then
    
        lr = .Cells(.Rows.Count, 1).End(xlUp).Row
        
        .Cells(lr + 1, 1) = mac_add
        
        MsgBox "初始化成功，歡迎使用!", vbInformation
        
        'ThisWorkbook.Close True

    End If
    
End With

End Sub

Function checkIsExist(ByVal mac_add As String)

checkIsExist = True

With Sheets("授權")
    
    Set rng = .Cells.Find(mac_add)
    
    If rng Is Nothing Then checkIsExist = False
    
End With

End Function

Sub hideCmd()

For Each sht In Sheets

    For Each shp In sht.Shapes
    
        'Debug.Print shp.OnAction
    
        If shp.OnAction <> "" Then shp.Visible = False
'
'        Select Case shp.Name
'
'        Case "Rounded Rectangle 5": shp.Visible = False
'        Case "Rounded Rectangle 6": shp.Visible = False
'
'        End Select

        
    Next

Next

End Sub

Sub showCmd()

For Each sht In Sheets

    For Each shp In sht.Shapes
    
        If shp.OnAction <> "" Then shp.Visible = True
    
'        Select Case shp.Name
'
'        Case "Rounded Rectangle 5": shp.Visible = True
'        Case "Rounded Rectangle 6": shp.Visible = True
'
'        End Select
        
    Next

Next

End Sub

Private Sub Class_Initialize()

'Sheets("授權").Visible = True

End Sub

Private Sub Class_Terminate()

'Sheets("授權").Visible = xlVeryHidden

End Sub
