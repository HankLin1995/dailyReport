Attribute VB_Name = "Main"
Function Eval(ByVal s As String)

Dim cal As String

For i = 1 To Len(s)

ch = mid(s, i, 1)

If IsNumeric(ch) Then '�P�_�O�_���Ʀr

cal = cal + ch

ElseIf ch = "(" Or ch = "[" Or ch = "{" Then  '�A��

cal = cal + "("

ElseIf ch = ")" Or ch = "]" Or ch = "}" Then    '�A��

cal = cal + ")"

ElseIf ch = "+" Or ch = "-" Or ch = "*" Or ch = "/" Or ch = "^" Then '�B���

cal = cal + ch

ElseIf ch = "." Then '��L����

cal = cal + ch

End If

Next

Eval = Application.Evaluate(cal)

End Function





