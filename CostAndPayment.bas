Attribute VB_Name = "CostAndPayment"
'估驗相關
'TODO:
'1.取得估驗日期
'2.取得估驗當日報表
'3.*判別是否有變更設計(取得報表內容為何)
'4.輸出至估驗表為初稿
'5.確認無誤後可傳至Cost_S儲存，設定為以前累計，計數估驗

Private costDay As Date

Sub setCostDay()

Dim costDate As Date

costDate = InputBox("請輸入估驗計價日期", , Format(Now(), "yyyy/mm/dd"))

Call FunctionModel.cmdGetReportIDByDate(costDate)

End Sub


