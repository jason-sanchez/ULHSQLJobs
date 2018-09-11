Imports System
Imports System.Data
Imports System.IO
Imports System.Data.SqlClient
Imports System.Net.Mail

Module Program
    ''Dim errorfilepath As String = "C:\Users\jsanchez\Documents\GitHub\ULHSQLJobs\ULHSQLJobs\ULHSQLJobs\Errors\"
    'Dim errorfilepath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory, "..\..\Errors\"))
    ''Dim completefilePath As String = "C:\Users\jsanchez\Documents\GitHub\ULHSQLJobs\ULHSQLJobs\ULHSQLJobs\CompletionLogs\"
    'Dim completefilePath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory, "..\..\CompletionLogs\"))

    Public Sub Main()

        'Dim errorfilepath As String = "C:\Users\jsanchez\Documents\GitHub\ULHSQLJobs\ULHSQLJobs\ULHSQLJobs\Errors\"
        Dim errorfilepath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory, "..\..\Errors\"))
        'Dim completefilePath As String = "C:\Users\jsanchez\Documents\GitHub\ULHSQLJobs\ULHSQLJobs\ULHSQLJobs\CompletionLogs\"
        Dim completefilePath As String = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory, "..\..\CompletionLogs\"))

        'Dim args() As String = Environment.GetCommandLineArgs()
        'errorfilepath = errorfilepath

        'completefilePath = completefilePath

        'args(0) = Full path of executing program with program name
        'args(1) - sql job
        'args(2) - server 
        'args(3) - database
        'args(4) - schema
        'Dim job As String = "smc_ULHSQLJobs_TEST"
        'Dim ip As String = "192.168.55.12\SQLEXPRESS"
        'Dim db As String = "ITW"
        'Dim schema As String = "dbo"
        'Dim ip As String = "10.48.64.5\SQLEXPRESS"
        'Dim db As String = "ITWTest"
        'Dim schema As String = "dbo"
        'If args.Length = 0 Then
        Dim errorfilename As String = String.Format("Error{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
            Dim errorfile = New StreamWriter(errorfilepath & errorfilename, True)
            errorfile.Write("Error")
            errorfile.Close()
        'Else
        'Dim errorfilename As String = String.Format("Error{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
        '    Dim errorfile = New StreamWriter(errorfilepath & errorfilename, True)
        '    errorfile.Write(args(0))
        '    errorfile.Close()


        'End If

        'Try
        '    'Using conn As New SqlConnection("server=" & args(1) & ";database=" & args(4) & ";uid=" & args(2) & ";pwd= " & args(3))
        '    Using conn As New SqlConnection("server=192.168.55.12\SQLEXPRESS;database=ITWLibraryMaster;uid=sysmax;pwd=sysmax")
        '        Dim objDBCommand As New SqlCommand
        '        With objDBCommand
        '            .Connection = conn
        '            .Connection.Open()
        '            .CommandText = args(1)
        '            .CommandType = CommandType.StoredProcedure
        '            .CommandTimeout = 1200 '10 min timeout
        '            .Parameters.AddWithValue("@IP", args(2))
        '            .Parameters.AddWithValue("@DB", args(3))
        '            .Parameters.AddWithValue("@Schema", args(4))

        '            '.CommandText = job
        '            '.CommandType = CommandType.StoredProcedure
        '            '.CommandTimeout = 1200 '10 min timeout
        '            '.Parameters.AddWithValue("@IP", ip)
        '            '.Parameters.AddWithValue("@DB", db)
        '            '.Parameters.AddWithValue("@Schema", schema)

        '            objDBCommand.ExecuteNonQuery()

        '        End With
        '    End Using

        '    'Log Completion
        '    logEvent("Job Complete", args(0) & " - Job Complete" & " - " & DateTime.Now, args(0))
        '    'logEvent("Job Complete", job & " - Job Complete" & " - " & DateTime.Now, job)

        'Catch ex As Exception
        '    logEvent("Error", DateTime.Now & " - " & ex.ToString(), args(0))
        '    sendemail(args(0) & " Error", ex.ToString())
        '    'logEvent("Error", DateTime.Now & " - " & ex.ToString(), job)
        '    'sendemail(job & " Error", ex.ToString())
        'End Try

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

    'Private Sub logEvent(ByVal type As String, ByVal mess As String, ByVal title As String)

    '    If type = "Error" Then
    '        System.IO.Directory.CreateDirectory(errorfilepath)
    '        Dim errorfilename As String = String.Format(title & " Error{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
    '        Dim errorfile = New StreamWriter(errorfilepath & errorfilename, True)
    '        errorfile.Write(mess)
    '        errorfile.Close()

    '    ElseIf type = "Email Error" Then
    '        System.IO.Directory.CreateDirectory(errorfilepath)
    '        Dim errorfilename As String = String.Format("EmailError{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
    '        Dim errorfile = New StreamWriter(errorfilepath & errorfilename, True)
    '        errorfile.Write(mess)
    '        errorfile.Close()

    '    ElseIf type = "Email Success" Then
    '        System.IO.Directory.CreateDirectory(completefilePath)
    '        Dim completefilename As String = String.Format("EmailSent{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
    '        Dim completefile = New StreamWriter(completefilePath & completefilename, True)
    '        completefile.Write(mess)
    '        completefile.Close()
    '    Else
    '        System.IO.Directory.CreateDirectory(completefilePath)
    '        Dim completefilename As String = String.Format(title & " Complete{0:yyyyMMdd_HH-mm-ss}.txt", Date.Now)
    '        Dim completefile = New StreamWriter(completefilePath & completefilename, True)
    '        completefile.Write(mess)
    '        completefile.Close()

    '    End If
    'End Sub


End Module

