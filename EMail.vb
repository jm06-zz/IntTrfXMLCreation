Imports System.Net.Mail
Public Class EMail

    Public sUsername As String
    Public sFromAddress As String
    Public sPassword As String
    Public sServer As String
    Public iPort As Integer
    Public bEnableSSL As Boolean
    Public sSubject As String
    Public sBody As String
    Public Recipients As New List(Of String)

    Function SendEmail(ByVal Recipients As List(Of String),
                      ByVal FromAddress As String,
                      ByVal Subject As String,
                      ByVal Body As String,
                      ByVal UserName As String,
                      ByVal Password As String,
                      Optional ByVal Server As String = "hqex3.lan.cafod.com",
                      Optional ByVal Port As Integer = 25,
                      Optional ByVal Attachments As List(Of String) = Nothing) As String
        Dim Email As New MailMessage()
        Try
            Dim SMTPServer As New SmtpClient
            For Each Attachment As String In Attachments
                Email.Attachments.Add(New Attachment(Attachment))
            Next
            Email.From = New MailAddress(FromAddress)
            For Each Recipient As String In Recipients
                Email.To.Add(Recipient)
            Next
            Email.Subject = Subject
            Email.Body = Body
            SMTPServer.Host = sServer
            SMTPServer.Port = iPort
            SMTPServer.Credentials = New System.Net.NetworkCredential(sUsername, sPassword)
            SMTPServer.EnableSsl = bEnableSSL
            SMTPServer.Send(Email)
            Email.Dispose()
            Return "Email to " & Recipients(0) & " from " & FromAddress & " was sent."
        Catch ex As SmtpException
            Email.Dispose()
            Return "Sending Email Failed. Smtp Error."
        Catch ex As ArgumentOutOfRangeException
            Email.Dispose()
            Return "Sending Email Failed. Check Port Number."
        Catch Ex As InvalidOperationException
            Email.Dispose()
            Return "Sending Email Failed. Check Port Number."
        End Try
    End Function



    Sub SetConnectionSettings()

        Dim sqCon As New SqlClient.SqlConnection(TfrFunctions.SQL_Connectionstring)
        Dim sqCmd As New SqlClient.SqlCommand
        Dim sdrRow As SqlClient.SqlDataReader


        ' see if T8 code mapping is present in the transfer mapping table
        sqCmd.Connection = sqCon
        sqCon.Open()

        sqCmd.CommandText = "SELECT Username, Password, Port, Server, EnableSSL, FromAddress FROM TFR_EMAIL_CONFIG WHERE Test = '" & sState & "'"

        sdrRow = sqCmd.ExecuteReader

        If sdrRow.Read() Then
            sUsername = sdrRow.GetValue(sdrRow.GetOrdinal("Username"))
            sPassword = sdrRow.GetValue(sdrRow.GetOrdinal("Password"))
            ' sPassword = ""
            iPort = CInt(sdrRow.GetValue(sdrRow.GetOrdinal("Port")))
            sServer = sdrRow.GetValue(sdrRow.GetOrdinal("Server"))
            bEnableSSL = Convert.ToBoolean(sdrRow.GetValue(sdrRow.GetOrdinal("EnableSSL")))
            ' bEnableSSL = False
            sFromAddress = sdrRow.GetValue(sdrRow.GetOrdinal("FromAddress"))
        Else
            MsgBox("No email config settings could be found!")
        End If
        sdrRow.Close()

        ' set recipient addresses according to email
        sqCmd.CommandText = "SELECT  * FROM TFR_ERROR_MAIL WHERE Test = '" & sState & "' AND Active = 'active'"


        sdrRow = sqCmd.ExecuteReader

        While sdrRow.Read()
            Recipients.Add(sdrRow.GetValue(sdrRow.GetOrdinal("EMail")))
        End While
        sqCon.Close()
    End Sub

    Sub SendErrorMail()
        Dim Attachments As New List(Of String)
        Dim sResponse As String

        sResponse = SendEmail(Recipients, sFromAddress, sSubject, sBody, sUsername, sPassword, sServer, iPort, Attachments)
        Console.WriteLine(sResponse)

#If DEBUG Then
        DirLogAppend.Log("SendErrorMail() " & sResponse)
#End If

    End Sub

End Class

