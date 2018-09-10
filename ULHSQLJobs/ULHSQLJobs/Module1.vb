Imports System
Imports System.Data
Imports System.IO
Imports System.Data.SqlClient
Imports System.Net.Mail

Module Program
    Dim errorfilepath As String = "D:\Errors"
    Dim completefilePath As String = "D:\CompletionLogs"

    Sub Main(args As String())

        'args(0) - sql job
        'args(1) - server 
        'args(2) - database
        'args(3) - schema

        Try
            'Using conn As New SqlConnection("server=" & args(1) & ";database=" & args(4) & ";uid=" & args(2) & ";pwd= " & args(3))
            Using conn As New SqlConnection("server=192.168.55.12\SQLEXPRESS;database=ITWLibraryMaster;uid=sysmax;pwd=sysmax")
                Dim objDBCommand As New SqlCommand
                With objDBCommand
                    .Connection = conn
                    .Connection.Open()
                    .CommandText = args(0)
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 1200 '10 min timeout
                    .Parameters.AddWithValue("@IP", args(1))
                    .Parameters.AddWithValue("@DB", args(2))
                    .Parameters.AddWithValue("@Schema", args(3))

                    objDBCommand.ExecuteNonQuery()

                End With
            End Using

            'Log Completion
            logEvent("Job Complete", args(0) & " - Job Complete" & " - " & DateTime.Now, args(0))

        Catch ex As Exception
            logEvent("Error", DateTime.Now & " - " & ex.ToString(), args(0))
            sendemail(args(0) & " Error", ex.ToString())
        End Try

    End Sub

    Private Sub sendemail(ByVal subject As String, ByVal mess As String)

        Dim SMTP As New SmtpClient("smtp.gmail.com")
        Dim message As New MailMessage("j0sanc02@gmail.com", "j0sanc02@gmail.com")
        Dim otherRecipient As String = "jsanchez@systemaxcorp.com"
        Dim username As String = "jsanchez@systemaxcorp.com"
        Dim password As String = "sanchez@819"
        Dim port As Integer = 587

        SMTP.EnableSsl = True

        SMTP.Credentials = New Net.NetworkCredential(username, password)
        SMTP.Port = port

        message.Subject = subject
        message.Body = mess

        message.To.Clear()
        message.To.Add(New MailAddress(otherRecipient))

        Try
            SMTP.Send(message)
            logEvent("Email Success", "Email Sent " & DateTime.Now, "")
        Catch ex As Exception
            logEvent("Email Error", DateTime.Now & " - " & ex.ToString(), "")
        End Try


    End Sub

    Private Sub logEvent(ByVal type As String, ByVal mess As String, ByVal title As String)

        If type = "Error" Then
            System.IO.Directory.CreateDirectory(errorfilepath)
            Dim errorfilename As String = String.Format(title & " Error{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
            Dim errorfile = New StreamWriter(errorfilepath & errorfilename, True)
            errorfile.Write(mess)
            errorfile.Close()

        ElseIf type = "Email Error" Then
            System.IO.Directory.CreateDirectory(errorfilepath)
            Dim errorfilename As String = String.Format("EmailError{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
            Dim errorfile = New StreamWriter(errorfilepath & errorfilename, True)
            errorfile.Write(mess)
            errorfile.Close()

        ElseIf type = "Email Success" Then
            System.IO.Directory.CreateDirectory(completefilePath)
            Dim completefilename As String = String.Format("EmailSent{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
            Dim completefile = New StreamWriter(completefilePath & completefilename, True)
            completefile.Write(mess)
            completefile.Close()
        Else
            System.IO.Directory.CreateDirectory(completefilePath)
            Dim completefilename As String = String.Format(title & " Complete{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
            Dim completefile = New StreamWriter(completefilePath & completefilename, True)
            completefile.Write(mess)
            completefile.Close()

        End If
    End Sub


End Module

