Sub SendEmail_Example1()

Dim EmailApp As Outlook.Application
Dim Source As String
Set EmailApp = New Outlook.Application

Dim EmailItem As Outlook.MailItem
Set EmailItem = EmailApp.CreateItem(olMailItem)

EmailItem.To = "xxxx@outlookmail.com"
EmailItem.CC = "xxxx@outlookmail.com"
EmailItem.BCC = "xxxx@outlookmail.com"
EmailItem.Subject = "Test Email From Excel VBA"
EmailItem.HTMLBody = "Hi," & vbNewLine & vbNewLine & "This chnaged is my first email from Excel" & _
vbNewLine & vbNewLine & _
"Regards," & vbNewLine & _
"VBA Coder"
Source = ThisWorkbook.FullName
EmailItem.Attachments.Add Source

EmailItem.Send

End Sub
