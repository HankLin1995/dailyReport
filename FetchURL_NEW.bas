Attribute VB_Name = "FetchURL_NEW"
Sub FetchURL_Main()

MAC_ADDRESS = getMacAddress

Debug.Print MAC_ADDRESS

If MAC_ADDRESS = "" Then
    Debug.Print "�нT�{�����O�_���s�W����~!"
Else

    Status = AccessStatus(MAC_ADDRESS)

    myChoose = Split(Status, ",")

    If myChoose(0) <> "PASS" Then

        MsgBox "�z�������w�L�k�~��ϥ�!", vbCritical
        
        frm_EndQRCode.Show
    
    Else
    
        Call ShowDialoge(myChoose(3))
        Call ShowDialoge(myChoose(2))  ', myChoose(3))

        '==========1206Add for ���s���� ============
        
        Dim o As New clsUserInformation
        o.showCmd

    End If

End If

End Sub

Sub ShowDialoge(ByVal s1 As String) ', ByVal s2 As String)

If s1 = "" Then Exit Sub

tmp = Split(s1, ":")

frmMSG.Label1.caption = tmp(0) ' caption
frmMSG.TextBox1.Value = tmp(1) ' s1
frmMSG.Show

End Sub

Function AccessStatus(ByVal mac_add As String) As String

KEEPACCESS:

Dim o As New clsFetchURL
Dim bIsClientSigned As Boolean

myURL = o.CreateURL("Access", mac_add)
Status = o.ExecHTTP(myURL)

myChoose = Split(Status, ",")

Select Case myChoose(0) 'Status

Case "PASS"

    Debug.Print "���ҳq�L!"
    
Case "NOT_FOUND"

    Debug.Print "�䤣���Ʈw���A�������Ǹ�"

    bIsClientSigned = IsClientSigned(mac_add) '�i����U���ըæ^�ǵ��G
    
    If bIsCliendSigned = False Then GoTo KEEPACCESS

Case "ARRIVED":

Case Else

    Debug.Print Status

End Select

AccessStatus = Status

End Function

Function IsClientSigned(ByVal mac_add As String) As Boolean

'�i����U�æ^�ǵ��U���A
'1.�w���U:True
'2.���U�q�L:False

IsClientSigned = False

Dim o As New clsFetchURL

myURL = o.CreateURL("Sign", mac_add)

If o.ExecHTTP(myURL) = "signed" Then
    MsgBox "�ӹq���w�g�Q���U�L�F!", vbCritical
    IsClientSigned = True
Else

myURL = o.CreateURL("SignDetail", mac_add, , , myMail)
Call o.ExecHTTP(myURL)

End If

End Function

'======�o�̽T�w���|��==========

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



