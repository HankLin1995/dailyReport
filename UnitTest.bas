Attribute VB_Name = "UnitTest"
Sub test_getRowsByUser2()

Dim o As New clsMyfunction

mydate = Format(CDate("2023/6/2"), "yyyy/mm/dd(aaa)")

Set coll = o.getRowsByUser2("Diary", mydate, 1, "������")

Debug.Assert coll(1) = 9

End Sub

Sub test_getPAYCount()

Dim myFunc As New clsMyfunction

Set coll_rows = myFunc.getUniqueItems("PAY_EX", 2, , "������")

Debug.Assert coll_rows.count + 1 = 1

End Sub

Sub test_exportSheets()

Dim o As New clsPrintOut
Dim f As String
f = Application.GetSaveAsFilename(InitialFileName:="�Ħ�����", FileFilter:="Excel Files (*.xlsx), *.xlsx")
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

Debug.Assert coll.count = 1

Set coll = f.getUniqueItems("PAY_EX", 2, , "������")

Debug.Assert coll.count = 1

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

MixName = "�G��"

result = o.getDetailUnitByMixName(MixName)

Debug.Assert result = "�y"
    
End Function

Function test_IsNumOnlyOne()

item_name = "���ҫO�@�A�o�󪫲M�z"

result = IsNumOnlyOne(item_name)

Debug.Assert result = True

item_name = "���s�Ҫ�"

result = IsNumOnlyOne(item_name)

Debug.Assert result = False

End Function







