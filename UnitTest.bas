Attribute VB_Name = "UnitTest"
Sub unittest_getRecLocInvolvedPrompt()

rec_loc = "0+350~0+3880" '表單新增
item_loc = "0+350~0+370" '已經紀錄

Dim f As New clsMyfunction

f.SplitAllLocs (rec_loc)

Debug.Print f.IsRecLocPass(rec_loc, item_loc)

End Sub

Sub test_IsTranLocContainALLIneed()

Dim f As New clsMyfunction

Debug.Assert f.IsNumericWithPlusAndParentheses("0+000.5~0+001") = True
Debug.Assert f.IsNumericWithPlusAndParentheses("dd+0+0") = False
Debug.Assert f.IsNumericWithPlusAndParentheses("0+000(上)") = True
Debug.Assert f.IsNumericWithPlusAndParentheses("0+000(其他)") = False

End Sub

Sub test_getMixLocPrompt_REC()

Dim o As New clsRecord
Dim f As New clsMyfunction


RecLocation = "2+350~2+390" '.txtWhere
RecItem = "土方工作，挖方" ' .cboItem
RecCanal = "單期一號"

'紀錄內容:1+340~1+380

For Each my_loc In f.SplitAllLocs(RecLocation)

    err_msg = o.getMixLocPrompt_REC(RecItem, my_loc, RecCanal)

    If err_msg <> "" Then p1 = p1 & err_msg & vbNewLine

Next

If p1 <> "" Then MsgBox p1, vbCritical

End Sub

Sub test_getMixLocPrompt_MIX()

Dim o As New clsRecord
Dim f As New clsMyfunction
Dim RecLocation  As String

'recDate = .txtDay
'RecChannelName = .cboChannel
RecLocation = "0+390~0+640" '.txtWhere
RecItem = "A-A',右牆" ' .cboItem
'If .txtAmount <> "" Then RecAmount = .txtAmount

If f.IsNumericWithPlusAndParentheses(RecLocation) = True Then

For Each my_loc In f.Spl
itAllLocs (RecLocation)

    p1 = p1 & o.getMixLocPrompt_MIX(RecItem, my_loc) & vbNewLine
    
Next

If p1 <> "" Then MsgBox p1, vbCritical
 
End If
 
End Sub

Sub test_IsMixNameUsed()

Dim o As New clsMixData

Debug.Assert o.IsMixNameUsed("C-C',右牆") = False

End Sub

Sub test_getNextNotBlankRow()

r = 71
lr = 72

Debug.Assert getNextNotBlankRow(r, lr) = 73

End Sub

Sub test_getRowsByUser2()

Dim o As New clsMyfunction

mydate = Format(CDate("2023/6/2"), "yyyy/mm/dd(aaa)")

Set coll = o.getRowsByUser2("Diary", mydate, 1, "報表日期")

Debug.Assert coll(1) = 9

End Sub

Sub test_getPAYCount()

Dim myFunc As New clsMyfunction

Set coll_rows = myFunc.getUniqueItems("PAY_EX", 2, , "估驗日期")

Debug.Assert coll_rows.Count + 1 = 1

End Sub

Sub test_exportSheets()

Dim o As New clsPrintOut
Dim f As String
f = Application.GetSaveAsFilename(initialFilename:="第次估驗", FileFilter:="Excel Files (*.xls), *.xls")
If f = "False" Then f = ""

Debug.Assert f = ""

End Sub

Sub test_CV()

Dim f As New clsMyfunction

Debug.Assert f.ConvertToLetter("6") = "F"

End Sub

Sub test_getUniqueItems()

Dim f As New clsMyfunction

Set coll = f.getUniqueItems("PAY_EX", 2, "F")

Debug.Assert coll.Count = 1

Set coll = f.getUniqueItems("PAY_EX", 2, , "估驗日期")

Debug.Assert coll.Count = 1

End Sub

Function test_IsPaydateLater()

Dim o As New clsPay

Debug.Assert o.IsPayDateLater(CDate("2023/7/16")) = False

End Function

Function test_IsRecDateInDiary()

Set shtDiary = Sheets("Diary")

ReportDay = "2023/9/23" 'Sheets("Report").Range("C2")

Debug.Print ReportDay

key_value = Format(CDate(ReportDay), "yyyy/mm/dd(aaa)")
Set rng = shtDiary.Columns("B").Find(what:=key_value, LookIn:=xlValues)

Debug.Assert rng Is Nothing

End Function

Function test_getDetailUnitByMixName()

'Arrange>>Act>>Assert

Dim o As New clsRecord

MixName = "矮堰"

result = o.getDetailUnitByMixName(MixName)

Debug.Assert result = "座"
    
End Function

Function test_IsNumOnlyOne()

item_name = "環境保護，廢棄物清理"

result = IsNumOnlyOne(item_name)

Debug.Assert result = True

item_name = "鋼製模版"

result = IsNumOnlyOne(item_name)

Debug.Assert result = False

End Function







