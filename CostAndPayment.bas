Attribute VB_Name = "CostAndPayment"
'�������
'TODO:
'1.���o������
'2.���o���������
'3.*�P�O�O�_���ܧ�]�p(���o�����e����)
'4.��X�ܦ������Z
'5.�T�{�L�~��i�Ǧ�Cost_S�x�s�A�]�w���H�e�֭p�A�p�Ʀ���

Private costDay As Date

Sub setCostDay()

Dim costDate As Date

costDate = InputBox("�п�J����p�����", , Format(Now(), "yyyy/mm/dd"))

Call FunctionModel.cmdGetReportIDByDate(costDate)

End Sub


