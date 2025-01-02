Attribute VB_Name = "tranFunction"

Sub Test()

strLoc = "28+500~28+525" '�B28+550.9~28+590"

If strLoc Like "*�B*" Then

    loc_tmp = Split(strLoc, "�B")
    
    For Each it In loc_tmp
    
        sumL = sumL + calcLoc(it)
        
    Next

Else

    sumL = calcLoc(strLoc)

End If

Debug.Print sumL

End Sub

Sub getSLocAndELoc(ByVal strLocation As String, ByRef sloc, ByRef eloc)

'If Not strLocation Like "*~*" Then MsgBox ("�����\����:" & vbNewLine & strLocation & vbNewLine & "�ШϥΡu~�v�i�����" & vbNewLine & "EX:0+000~0+100"), vbCritical: End

loc_split = Split(strLocation, "~")

sloc = TranLoc(loc_split(0))
eloc = TranLoc(loc_split(1))

End Sub

Function calcLoc(ByVal strLocation As String) 'only for "0+000~0+000"

On Error GoTo ERRORHANDLE

strLocation = Replace(strLocation, "��", "~")

loc_split = Split(strLocation, "~")

sloc = loc_split(0)
eloc = loc_split(1)
'
calcLoc = TranLoc(eloc) - TranLoc(sloc)

Exit Function

ERRORHANDLE:

MsgBox "�榡�ݭn��0+000~0+000!", vbCritical

End Function

Function TranLoc(ByVal Data As String) As Double

'�θ����A�ন�i�p�⤧�θ�

tmp = Split(Data, "+")

If UBound(tmp) = -1 Or Data = "" Then Exit Function ' TranLoc = CDbl(Data): Exit Function

tloc = tmp(0) '�d���
dloc = tmp(1)

If dloc Like "*(*" Then

    tmp2 = Split(dloc, "(")

    If tmp2(0) Like "*.*" Then

        tmp3 = Split(tmp2(0), ".")
        dloc = tmp3(0) + tmp3(1) / 10
    
    Else
    
        dloc = tmp2(0)
    
    End If
    
End If

For i = 1 To Len(tloc)

    ch = mid(tloc, i, 1)
    If IsNumeric(ch) Then ref = ref & ch

Next

TranLoc = CDbl(ref) * 1000 + CDbl(dloc)
    
End Function
