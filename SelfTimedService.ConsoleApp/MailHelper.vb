
''' <summary>
''' Simple helper class to handler mailing.
''' </summary>
Public Class MailHelper

    Public Property Usr As String
    Public Property Pwd As String
    Public Property AppName As String
    Public Property MailTo As String()
    Public Property MailSmtpClient As String = "smtp.gmail.com"
    Public Property EnableSSL As Boolean


    Public Sub New(MailUser As String, MailPassword As String, MailTo As String(), Optional SmtpClient As String = Nothing)
        Me.AppName = AppName
        Me.Usr = MailUser
        Me.Pwd = MailPassword
        Me.MailTo = MailTo
        Me.MailSmtpClient = If(SmtpClient, MailSmtpClient)
    End Sub

    ''' <summary>
    ''' The main functionality of the MailHelper. It'll try to default every parameter but the body. If it fails to get a User, Password or MailTo, it'll throw.
    ''' </summary>
    ''' <param name="MailBody">The body</param>
    ''' <param name="MailTitle">If not provided, it will default to the AppName property.</param>
    ''' <param name="MailTo">If not provided, it will try to default to the MailTo provided in the constructor (if any). Lastly, it'll try to use "MailTo" from the App.config</param>
    ''' <param name="EnableSSL">It'll default to False.</param>
    ''' <param name="SmtpClient">It'll default to "smtp.gmail.com".</param>
    Public Sub Send(MailBody As String, Optional MailTitle As String = Nothing,
                    Optional MailTo As String() = Nothing, Optional EnableSSL As Boolean? = False,
                    Optional SmtpClient As String = Nothing)

        Dim Usr As String = Me.Usr
        Dim Pwd As String = Me.Pwd
        Dim Title As String = MailTitle
        Dim TheSmtpClient = SmtpClient
        MailTo = If(MailTo, Me.MailTo)

        If (ShouldThrow(Usr, Pwd, MailTo)) Then
            Throw New ArgumentException("The mail needs a user, password and mailto.")
        End If

        EmailUtil.Send(Usr, MailTo, If(Title, AppName), MailBody, Usr, Pwd, TheSmtpClient, EnableSSL)

    End Sub

    Function ShouldThrow(Usr As String, Pwd As String, MailTo As String()) As Boolean

        Return String.IsNullOrWhiteSpace(Usr) Or String.IsNullOrWhiteSpace(Pwd) _
        Or MailTo Is Nothing Or MailTo.All(Function(x) String.IsNullOrWhiteSpace(x))

    End Function

End Class