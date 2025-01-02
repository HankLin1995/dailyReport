Attribute VB_Name = "mail_test"
Sub SendEmail(ByVal RecipientEmail As String, ByVal AttachmentPath As String)
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    'Dim RecipientEmail As String
    Dim EmailSubject As String
    Dim EmailBody As String
    'Dim AttachmentPath As String
    
    ' 收件者的郵件地址
   ' RecipientEmail = "apple84026113@gmail.com"
    
    ' 郵件主題
    EmailSubject = "檢驗停留點申請單"
    
    ' 郵件內容
    EmailBody = "主辦監造您好，這是本期工程的檢驗停留點申請單，敬請查悉。"
    
    ' 附件的文件路徑
    'AttachmentPath = ThisWorkbook.Path & "\抽查表Output\EN-1.xls" '"C:\Path\To\Your\File.pdf"
    
    ' 創建一個Outlook應用程式對象
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0) ' 0表示郵件項目
    
    ' 設置郵件屬性
    With OutlookMail
        .To = RecipientEmail
        .Subject = EmailSubject
        .Body = EmailBody
        .Attachments.Add AttachmentPath ' 添加附件
        .Send ' 發送郵件
    End With
    
    ' 釋放對象
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub

