Attribute VB_Name = "Main"
Function Eval(ByVal s As String)

Dim cal As String

For i = 1 To Len(s)

ch = mid(s, i, 1)

If IsNumeric(ch) Then '判斷是否為數字

cal = cal + ch

ElseIf ch = "(" Or ch = "[" Or ch = "{" Then  '括弧

cal = cal + "("

ElseIf ch = ")" Or ch = "]" Or ch = "}" Then    '括弧

cal = cal + ")"

ElseIf ch = "+" Or ch = "-" Or ch = "*" Or ch = "/" Or ch = "^" Then '運算符

cal = cal + ch

ElseIf ch = "." Then '其他項目

cal = cal + ch

End If

Next

Eval = Application.Evaluate(cal)

End Function





