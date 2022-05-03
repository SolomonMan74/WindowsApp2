Imports Microsoft.VisualBasic
Imports System.Net.Mail
Imports System.Configuration
Imports System.Data.SqlClient
Public Class Form1

    Public Function GetCodeValue(ByVal code As String)

        Dim ReturnVal As String = ""
        If code = "MailServer" Then
            ReturnVal = "tmd-exh04.ad.tmdinc.com"
        Else
            ReturnVal = "noreply@grammer.com"
        End If

        Return ReturnVal

    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Public Function SendEmail(ByVal MailSubject As String, ByVal MailTo As String, ByVal MailText As String, ByVal Optional HTMLMail As Boolean = False, Optional ByVal MailFrom As String = "XXXXX") As Boolean
        Dim Success As Boolean = False
        'create the mail message
        Dim mail As New MailMessage()
        Dim MailServer As String = GetCodeValue("MAILSERVER")
        Dim FromEmail As String = GetCodeValue("MAILFROM")
        If MailFrom = "XXXXX" Then
            MailFrom = FromEmail
        End If
        'set the addresses
        mail.From = New MailAddress(MailFrom)
        mail.[To].Add(MailTo)
        'set the subject
        mail.Subject = MailSubject
        mail.IsBodyHtml = HTMLMail
        'set the content body
        mail.Body = MailText

        'set the server
        Dim smtp As New SmtpClient(MailServer)

        'Send the eMail
        smtp.Send(mail)
        Return Success
    End Function

    Private Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click
        Dim OutGoEmailAddress As String = "chris.gray@grammer.com" '
        'Dim OutGoEmailAddress As String = "COI-IT-Dev-Notifications@grammer.com"
        SendEmail("wcaUpdateActiveDirectoryFromQuest - Import to Active Directory", OutGoEmailAddress, "Import Successful", False)

    End Sub
End Class
