Attribute VB_Name = "UnitTest"
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







