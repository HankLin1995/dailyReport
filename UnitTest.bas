Attribute VB_Name = "UnitTest"
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







