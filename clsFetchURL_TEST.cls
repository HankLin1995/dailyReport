VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFetchURL_TEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'HankLin edited @ 20210823

Dim XMLHTTP As Object
Dim DOM As Object
Private URL As String

'Const URL As String = "https://script.google.com/macros/s/AKfycbzNvQOSQoBc7h3bnyYTmiw_zy8ijUl-IiDpkl-Zw21mb59exRbj_mp4p323giDPcOyA/exec"

Sub Class_Initialize()

Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
Set DOM = CreateObject("Htmlfile")

URL = Sheets("Test").Range("E1")

End Sub

Function CreateURL(ByVal test_name As String, ByVal test_num As Integer)


test_name = getUnicode(test_name)
test_num = CStr(test_num)

CreateURL = URL & "?test_name=" & test_name & "&test_num=" & test_num
                    
End Function

Function ExecHTTP(ByVal myURL As String)

'On Error GoTo ERRORHANDLE

With XMLHTTP

    .Open "GET", myURL, False
    .Send

    
    If .Status = 200 Then
    
        'tmp = Split(.responsetext, "***")
        
        'If UBound(tmp) <> 0 Then
        
          '  ExecHTTP = tmp(1)
        
        'End If
        
    Else
'ERRORHANDLE:
       ' MsgBox "Something ERROR!", vbCritical
       ' End

    End If
    
End With

End Function
    
Private Function getUnicode(s)

For i = 1 To Len(s)

    myUnicode = myUnicode & "," & CharToUnicode(mid(s, i, 1))

Next

getUnicode = mid(myUnicode, 2)

End Function

    
Private Function CharToUnicode(strChar As String)

    'https://stackoverflow.com/questions/4006246/getting-the-unicode-value-of-a-char-in-vb/34780532

    Dim lngUnicode As Long

    lngUnicode = AscW(strChar)

    If lngUnicode < 0 Then
        lngUnicode = 65536 + lngUnicode
    End If

    CharToUnicode = lngUnicode
    
End Function






