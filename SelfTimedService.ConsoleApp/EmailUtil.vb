Imports System.Net
Imports System.Net.Mail

Public Class EmailUtil
    Shared DefaultClient As String = "smtp.gmail.com"

    Public Shared Sub Send(ByVal From As String, ByVal [To] As String, ByVal Subject As String, ByVal Body As String, ByVal Username As String, ByVal Password As String, ByVal Optional SmtpClient As String = Nothing, ByVal Optional EnableSSL As Boolean = True)
        Dim Mail As MailMessage = CreateMail(From, New String() {[To]}, Subject, Body)
        Dim Credential As NetworkCredential = CreateCredential(Username, Password)
        SmtpDispatch(Credential, Mail, If(SmtpClient, DefaultClient), EnableSSL)
    End Sub

    Public Shared Sub Send(ByVal From As String, ByVal [To] As String(), ByVal Subject As String, ByVal Body As String, ByVal Username As String, ByVal Password As String, ByVal Optional SmtpClient As String = Nothing, ByVal Optional EnableSSL As Boolean = True)
        Dim Mail As MailMessage = CreateMail(From, [To], Subject, Body)
        Dim Credential As NetworkCredential = CreateCredential(Username, Password)
        SmtpDispatch(Credential, Mail, If(SmtpClient, DefaultClient), EnableSSL)
    End Sub

    Public Shared Sub Send(ByVal Message As MailMessage, ByVal Credentials As NetworkCredential, attachments As IEnumerable(Of Attachment), ByVal Optional SmtpClient As String = Nothing, ByVal Optional EnableSSL As Boolean = True)
        Message.Attachments.AddRange(attachments)
        SmtpDispatch(Credentials, Message, If(SmtpClient, DefaultClient), EnableSSL)
    End Sub

    Public Shared Sub Send(ByVal Message As MailMessage, ByVal Credentials As NetworkCredential, ByVal Optional SmtpClient As String = Nothing, ByVal Optional EnableSSL As Boolean = True)
        SmtpDispatch(Credentials, Message, If(SmtpClient, DefaultClient), EnableSSL)
    End Sub

    Private Shared Sub SmtpDispatch(ByVal Credentials As NetworkCredential, ByVal Message As MailMessage, ByVal SmtpClient As String, ByVal EnableSSL As Boolean)
        Dim SmtpServer As SmtpClient = New SmtpClient(SmtpClient)
        SmtpServer.Port = 587
        SmtpServer.Credentials = Credentials
        SmtpServer.EnableSsl = True
        SmtpServer.Send(Message)
    End Sub

    Private Shared Function CreateMail(ByVal From As String, ByVal [To] As String(), ByVal Subject As String, ByVal Body As String) As MailMessage
        Dim mail As MailMessage = New MailMessage()
        mail.From = New MailAddress(From)
        mail.[To].AddRange([To])
        mail.Subject = Subject
        mail.Body = Body
        Return mail
    End Function

    Private Shared Function CreateCredential(ByVal Username As String, ByVal Password As String) As NetworkCredential
        Return New NetworkCredential(Username, Password)
    End Function
End Class

