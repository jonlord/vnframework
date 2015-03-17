Imports System.IO
Imports System.Net.Mail

''' <summary>
''' Email wrapper class to send mails using SMTPc
''' </summary>
Public Class email
    ''' <summary>
    ''' SMTP server username
    ''' </summary>
    Private _smtpUserName As String
    ''' <summary>
    ''' Sender's display name
    ''' </summary>
    Private _smtpName As String
    ''' <summary>
    ''' SMTP server password
    ''' </summary>
    Private _smtpPassword As String
    ''' <summary>
    ''' SMTP server port
    ''' </summary>
    Private _smtpPort As Integer
    ''' <summary>
    ''' SMTP server address
    ''' </summary>
    Private _smtpServer As String

    ''' <summary>
    ''' Constructs an email object to send emails
    ''' </summary>
    ''' <param name="name">Sender's display name</param>
    ''' <param name="userName">SMTP server username</param>
    ''' <param name="password">SMTP server password</param>
    ''' <param name="port">SMTP server port</param>
    ''' <param name="server">SMTP server address</param>
    Sub New(ByVal name As String, ByVal userName As String, ByVal password As String, ByVal port As Integer, ByVal server As String)
        _smtpUserName = userName
        _smtpName = name
        _smtpPassword = password
        _smtpPort = port
        _smtpServer = server
    End Sub

    ''' <summary>
    ''' Sends an email to the specified list of address, allowing for an optional attachment
    ''' </summary>
    ''' <param name="addresses">Email's Recipient Address (separated by commas)</param>
    ''' <param name="subject">Email's subject</param>
    ''' <param name="content">Email's Content</param>
    ''' <param name="attachment">Path to attachment file</param>
    ''' <returns></returns>
    Function send(ByVal addresses As String, ByVal subject As String, ByVal content As String, Optional ByVal attachment As String = "") As Boolean
        Using SmtpServer As New SmtpClient() With {.Credentials = New Net.NetworkCredential(_smtpUserName, _smtpPassword), .Port = _smtpPort, .Host = _smtpServer, .EnableSsl = True}
            Using mail As New MailMessage()
                Dim addr() As String = addresses.Split(",")
                mail.IsBodyHtml = True
                mail.From = New MailAddress(_smtpUserName, _smtpName, Text.Encoding.UTF8)
                Dim i As Byte
                For i = 0 To addr.Length - 1
                    mail.To.Add(addr(i))
                Next
                If attachment <> "" Then
                    If Not File.Exists(attachment) Then
                        Throw New FileNotFoundException
                        Return False
                    Else
                        mail.Attachments.Add(New Attachment(attachment))
                    End If
                End If
                mail.Subject = subject
                mail.Body = content
                mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure
                mail.ReplyToList.Add(_smtpUserName)
                SmtpServer.Send(mail)
            End Using
        End Using
        Return True
    End Function
End Class