Attribute VB_Name = "UnitTest"

Sub unittest_Main()

unittest_getRecLocInvolvedPrompt
unittest_getExRecItem
unittest_IsTranLocContainALLIneed

unittest_getMixItemsByChname ("�g��p��2-5")

End Sub

Sub unittest_getCorrectPgs()

Dim o As New clsReport

Dim pgs_today As Double
Dim pgs_total As Double

rec_date = Sheets("Report").Range("C2")

Call o.getCorrectPgs(rec_date, pgs_today, pgs_total)

Debug.Assert pgs_today = 0
Debug.Assert pgs_total = 0.0111

End Sub


Sub unittest_getCustomOrder()

Dim myfunc As New clsMyfunction

Set collPropIndex = getSepIndexByChname("�g��p��2-5")

Set new_collPropIndex = myfunc.changeOrder(collPropIndex)

Debug.Assert new_collPropIndex.Count = 0

End Sub


Sub unittes_getPropByMixName()

Debug.Assert getPropByMixName("�p��3����0��100") = "����"

Set collPropIndex = tmp_code.test_getSepIndexByChname("�g��p��2-5")

myindex = tmp_code.getCollIndex(collPropIndex, "����")

Debug.Assert myindex = 1

End Sub

Sub unittest_getSepIndexByChname()

chname = "�g��p��2-5"

Set collPropIndex = tmp_code.test_getSepIndexByChname(chname)

Debug.Assert collPropIndex.Count > 0

End Sub

Sub unittest_getMixItemsByChname(chname)

Set coll = progress_plot.getMixItemsByChname(chname)

Debug.Assert coll.Count > 0

End Sub

Sub unittest_getRecLocInvolvedPrompt()

rec_loc = "0+350~0+388" '���s�W
item_loc = "0+350~0+370" '�w�g����

Dim f As New clsMyfunction

f.SplitAllLocs (rec_loc)

Debug.Assert f.IsRecLocPass(rec_loc, item_loc) = False

Debug.Print "getRecLocInvolvedPrompt...PASS"

End Sub


Sub unittest_getExRecItem()

Debug.Assert getExRecItem("�p��2-5�k��216��316") = "�p��2-5�j��216��316"
Debug.Assert getExRecItem("�p��2-5����216��316") = "�p��2-5�j��216��316"
Debug.Assert getExRecItem("�p��2-5�j��216��316") = "�p��2-5����216��316"
Debug.Assert getExRecItem("�p��2-5����216��316") = ""

Debug.Print "getExRecItem...PASS"

End Sub

Sub unittest_IsTranLocContainALLIneed()

Dim f As New clsMyfunction

Debug.Assert f.IsNumericWithPlusAndParentheses("0+000.5~0+001") = True
Debug.Assert f.IsNumericWithPlusAndParentheses("dd+0+0") = False
Debug.Assert f.IsNumericWithPlusAndParentheses("0+000(�W)") = True
Debug.Assert f.IsNumericWithPlusAndParentheses("0+000(��L)") = False

Debug.Print "IsTranLocContainALLIneed...PASS"

End Sub

Sub test_countCheckLists()

check_date = CDate("2023/8/2")

Dim o As New clsCheck

result = o.countCheckLists(check_date)

Debug.Assert result = 0

End Sub


Sub test_SplitAllLocs()

Dim f As New clsMyfunction

Set coll_locs = f.SplitAllLocs("0+000~0+100�B0+100~0+200")

For Each it In coll_locs

    Debug.Print it

Next

End Sub

Sub test_getMixLocPrompt_REC()

Dim o As New clsRecord
Dim f As New clsMyfunction


RecLocation = "2+350~2+390" '.txtWhere
RecItem = "�g��u�@�A����" ' .cboItem
RecCanal = "����@��"

'�������e:1+340~1+380

For Each my_loc In f.SplitAllLocs(RecLocation)

    err_msg = o.getMixLocPrompt_REC(RecItem, my_loc, RecCanal)

    If err_msg <> "" Then p1 = p1 & err_msg & vbNewLine

Next

If p1 <> "" Then MsgBox p1, vbCritical

End Sub


Sub test_getMixLocPrompt_MIX() '20241212�P�w�e�����جO�_���I�@�A����!

Dim o As New clsRecord
Dim f As New clsMyfunction
Dim RecLocation  As String

'recDate = .txtDay
'RecChannelName = .cboChannel
RecLocation = "0+305~0+312" '.txtWhere
RecItem = "�p��2-5�k��216��316" ' .cboItem

If RecItem = "" Then Exit Sub

Set coll_rec_locs = f.SplitAllLocs(RecLocation)
Set coll_rows = f.getRowsByUser2("Records", RecItem, 2, "�զX�u��")

With Sheets("Records")

    For Each r In coll_rows
    
        item_loc_origin = .Cells(r, "D")

        If f.IsNumericWithPlusAndParentheses(CStr(rec_loc)) = True Then
        
            Set coll_item_locs = f.SplitAllLocs(item_loc_origin)
        
            For Each rec_loc In coll_rec_locs
        
                For Each item_loc In coll_item_locs
            
                    If f.IsRecLocPass(rec_loc, item_loc) = False Then
                        getMixLocPrompt_MIX_prompt = "��" & r & "�C:�i" & item_loc & "�j�P������i" & rec_loc & "�j�Ĭ�!":
                    End If
                
                Next
            
            Next
            
        End If
        
    Next

End With

Debug.Assert getMixLocPrompt_MIX_prompt = ""
 
End Sub

Sub unittest_IsMixNameUsed()

Dim o As New clsMixData

Debug.Assert o.IsMixNameUsed("C-C',�k��") = False
Debug.Assert o.IsMixNameUsed("�p��2-5����716��1064") = True

End Sub

Sub test_getNextNotBlankRow()

r = 71
lr = 72

Debug.Assert getNextNotBlankRow(r, lr) = 73

End Sub

Sub test_getRowsByUser2()

Dim o As New clsMyfunction

mydate = Format(CDate("2023/6/2"), "yyyy/mm/dd(aaa)")

Set coll = o.getRowsByUser2("Diary", mydate, 1, "������")

Debug.Assert coll(1) = 9

End Sub

Sub test_getPAYCount()

Dim myfunc As New clsMyfunction

Set coll_rows = myfunc.getUniqueItems("PAY_EX", 2, , "������")

Debug.Assert coll_rows.Count + 1 = 1

End Sub

Sub test_exportSheets()

Dim o As New clsPrintOut
Dim f As String
f = Application.GetSaveAsFilename(initialFilename:="�Ħ�����", FileFilter:="Excel Files (*.xls), *.xls")
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

Set coll = f.getUniqueItems("PAY_EX", 2, , "������")

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







