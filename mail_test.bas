Attribute VB_Name = "mail_test"
Sub SendEmail(ByVal RecipientEmail As String, ByVal AttachmentPath As String)
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    'Dim RecipientEmail As String
    Dim EmailSubject As String
    Dim EmailBody As String
    'Dim AttachmentPath As String
    
    ' ����̪��l��a�}
   ' RecipientEmail = "apple84026113@gmail.com"
    
    ' �l��D�D
    EmailSubject = "���簱�d�I�ӽг�"
    
    ' �l�󤺮e
    EmailBody = "�D��ʳy�z�n�A�o�O�����u�{�����簱�d�I�ӽг�A�q�Ьd�x�C"
    
    ' ���󪺤����|
    'AttachmentPath = ThisWorkbook.Path & "\��d��Output\EN-1.xls" '"C:\Path\To\Your\File.pdf"
    
    ' �Ыؤ@��Outlook���ε{����H
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0) ' 0��ܶl�󶵥�
    
    ' �]�m�l���ݩ�
    With OutlookMail
        .To = RecipientEmail
        .Subject = EmailSubject
        .Body = EmailBody
        .Attachments.Add AttachmentPath ' �K�[����
        .Send ' �o�e�l��
    End With
    
    ' �����H
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub

