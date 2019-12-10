Imports System.Net.Mail
Imports System.Runtime.CompilerServices

Module MailExtensions
    <Extension()>
    Sub AddRange(ByVal attachment As AttachmentCollection, ParamArray attachments As Attachment())
        For Each a In attachments
            attachment.Add(a)
        Next
    End Sub

    <Extension()>
    Sub AddRange(ByVal AddressCollection As MailAddressCollection, ParamArray addresses As String())
        For Each a In addresses
            AddressCollection.Add(a)
        Next
    End Sub
End Module
