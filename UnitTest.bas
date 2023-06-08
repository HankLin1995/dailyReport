Attribute VB_Name = "UnitTest"
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







