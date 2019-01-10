Imports System
Imports System.Data
Imports System.IO
Imports System.Data.SqlClient
Imports System.Net.Mail

Module Program
    'Dim errorfilepath As String = "C:\Users\jsanchez\Documents\GitHub\ULHSQLJobs\ULHSQLJobs\ULHSQLJobs\Errors\"
    Dim errorfilepath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory, System.Configuration.ConfigurationManager.AppSettings("errorfilepath")))
    'Dim completefilePath As String = "C:\Users\jsanchez\Documents\GitHub\ULHSQLJobs\ULHSQLJobs\ULHSQLJobs\CompletionLogs\"
    Dim completefilePath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory, System.Configuration.ConfigurationManager.AppSettings("completefilepath")))
    Dim daysback As Integer = System.Configuration.ConfigurationManager.AppSettings("daysback")
    Dim job As String = ""

    Public Sub Main()

        ' get parameters. (0) is the executable.  
        Try
            Dim args() As String = Environment.GetCommandLineArgs()
            Dim dbstring As String = args(1).ToString()
            job = args(2).ToString()

            errorfilepath = errorfilepath & job & "\"
            completefilePath = completefilePath & job & "\"

            If (Not System.IO.Directory.Exists(errorfilepath)) Then
                System.IO.Directory.CreateDirectory(errorfilepath)
            End If

            If (Not System.IO.Directory.Exists(completefilePath)) Then
                System.IO.Directory.CreateDirectory(completefilePath)
            End If

            Using conn As New SqlConnection(dbstring)
                Dim objDBCommand As New SqlCommand
                With objDBCommand
                    .Connection = conn
                    .Connection.Open()
                    .CommandText = job
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 1200 '10 min timeout
                    objDBCommand.ExecuteNonQuery()

                    logEvent("Success", job & "Job Complete - " & DateTime.Now, job, completefilePath)

                End With

            End Using

        Catch ex As Exception
            ex.ToString()

            logEvent("Error", DateTime.Now & " - " & ex.ToString(), job, errorfilepath)
            'sendemail(args(0) & " Error", ex.ToString())
            'logEvent("Error", DateTime.Now & " - " & ex.ToString(), job)
            'sendemail(job & " Error", ex.ToString())

        Finally
            RemoveOldLogs(completefilePath)
            RemoveOldLogs(errorfilepath)
        End Try

    End Sub

    Private Sub RemoveOldLogs(ByVal filePath As String)

        Try
            For Each logfile As String In Directory.GetFiles(filePath)
                Dim diff As Int16 = DateDiff("d", FileDateTime(logfile), Now)
                If DateDiff("d", FileDateTime(logfile), Now) > daysback Then
                    My.Computer.FileSystem.DeleteFile(logfile)
                End If
            Next
        Catch ex As Exception
            ex.ToString()
            logEvent("Error", DateTime.Now & " - " & ex.ToString(), "Log Removal Error", errorfilepath)
        End Try


    End Sub

    'Private Sub sendemail(ByVal subject As String, ByVal mess As String)

    '    Dim SMTP As New SmtpClient("smtp.gmail.com")
    '    Dim message As New MailMessage("j0sanc02@gmail.com", "j0sanc02@gmail.com")
    '    Dim otherRecipient As String = "jsanchez@systemaxcorp.com"
    '    Dim username As String = "jsanchez@systemaxcorp.com"
    '    Dim password As String = "sanchez@819"
    '    Dim port As Integer = 587

    '    SMTP.EnableSsl = True

    '    SMTP.Credentials = New Net.NetworkCredential(username, password)
    '    SMTP.Port = port

    '    message.Subject = subject
    '    message.Body = mess

    '    message.To.Clear()
    '    message.To.Add(New MailAddress(otherRecipient))

    '    Try
    '        SMTP.Send(message)
    '        logEvent("Email Success", "Email Sent " & DateTime.Now, "")
    '    Catch ex As Exception
    '        logEvent("Email Error", DateTime.Now & " - " & ex.ToString(), "")
    '    End Try


    'End Sub

    Private Sub logEvent(ByVal type As String, ByVal mess As String, ByVal title As String, ByVal filepath As String)

        If type = "Error" Then

            Dim errorfilename As String = String.Format(title & " Error{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
            Dim errorfile = New StreamWriter(filepath & errorfilename, True)
            errorfile.Write(mess)
            errorfile.Close()

            'ElseIf type = "Email Error" Then
            '    System.IO.Directory.CreateDirectory(errorfilepath)
            '    Dim errorfilename As String = String.Format("EmailError{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
            '    Dim errorfile = New StreamWriter(errorfilepath & errorfilename, True)
            '    errorfile.Write(mess)
            '    errorfile.Close()

            'ElseIf type = "Email Success" Then
            '    System.IO.Directory.CreateDirectory(completefilePath)
            '    Dim completefilename As String = String.Format("EmailSent{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
            '    Dim completefile = New StreamWriter(completefilePath & completefilename, True)
            '    completefile.Write(mess)
            '    completefile.Close()
        Else

            Dim completefilename As String = String.Format(title & " Complete - {0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
            Dim completefile = New StreamWriter(filepath & completefilename, True)
            completefile.Write(mess)
            completefile.Close()

        End If
    End Sub


End Module

