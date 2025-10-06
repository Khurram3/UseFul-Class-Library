Imports System.Data.OleDb
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Net.NetworkInformation
Imports System.Net.Sockets
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.Win32

Public Class ClsUseFulFunctions

    Public Function StringToNumber(ByVal mytext As String) As String
        Dim myChars() As Char = mytext.ToCharArray()
        For Each ch As Char In myChars
            If Char.IsDigit(ch) Or ch = "." Then
                mytext = mytext + (ch)
            End If
        Next
        StringToNumber = mytext
        Return StringToNumber
    End Function

    Public Function FilterNumbers(input As String) As String
        Dim regex As New Regex("\d+")
        Dim matches As MatchCollection = regex.Matches(input)
        Dim result As String = String.Join("", matches.Cast(Of Match).Select(Function(m) m.Value))
        Return result
    End Function
    Public Sub SpeckName(ByVal TextToSpeak As String)
        Dim SAPI
        SAPI = CreateObject("SAPI.spvoice")
        SAPI.Speak(TextToSpeak)
    End Sub
    Public Function RegistryDataGet(ByVal AppName As String, ByVal KeyName As String, ByVal KeyValue As String) As String
        Dim CU As RegistryKey = Registry.CurrentUser.CreateSubKey(AppName)
        With CU
            .OpenSubKey("Folder", True)
            KeyValue = .GetValue(KeyName)
            .Close()
        End With
        If KeyValue Is Nothing Then
            KeyValue = False
        End If
        Return KeyValue
    End Function
    Public Function RegistryDataGet(ByVal AppName As String, ByVal KeyName As String, ByVal KeyValue As String, ByVal SubKeyValue As Boolean) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("KeyName")
        dt.Columns.Add("KeyValue")
        Dim CU As RegistryKey = Registry.CurrentUser.CreateSubKey(AppName)
        With CU
            If SubKeyValue = False Then
                Dim subKeyNames As String() = .GetSubKeyNames()
                For Each subKeyName As String In subKeyNames
                    dt.Rows.Add(subKeyName, subKeyName)
                Next
            Else
                .OpenSubKey("Folder", True)
                For Each Val As String In .GetValueNames
                    dt.Rows.Add(Val, .GetValue(Val, Val))
                Next
            End If
            .Close()
        End With
        Return dt
    End Function
    Public Sub RegistryDataSet(ByVal AppName As String, ByVal KeyName As String, ByVal KeyValue As String)
        On Error Resume Next
        Dim CU As RegistryKey = Registry.CurrentUser.CreateSubKey(AppName)
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & AppName, "Folder", Nothing) Is Nothing Then
            With CU
                .OpenSubKey("Folder")
                .SetValue(KeyName, KeyValue)
                .Close()
            End With
        End If
    End Sub
    Public Sub RegistryDataDelete(ByVal AppName As String, ByVal KeyName As String, ByVal KeyValue As String)
        On Error Resume Next
        Dim CU As RegistryKey = Registry.CurrentUser.CreateSubKey(AppName)
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & AppName, "Folder", Nothing) Is Nothing Then
            With CU
                .OpenSubKey("Folder")
                .DeleteValue(KeyName)
                .Close()
            End With
        End If
    End Sub
    Public Sub RegistryDataDelete(ByVal AppName As String, ByVal KeyName As String, ByVal KeyValue As String, ByVal All As Boolean)
        On Error Resume Next
        Dim CU As RegistryKey = Registry.CurrentUser.CreateSubKey(AppName)
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\" & AppName, "Folder", Nothing) Is Nothing Then
            With CU
                .OpenSubKey("Folder")
                For Each val As String In .GetValueNames
                    CU.DeleteValue(val)
                Next
                .Close()
            End With
        End If
    End Sub
    Public Sub SaveLog(ByVal FilePath As String, ByVal FileName As String, ByVal FileData As String, ByVal FileType As String)
        On Error Resume Next
        Dim FileDate As String = Now.ToString("yyyy-MM-dd")
        Dim FullPath As String = Path.Combine(FilePath, FileDate & " " & FileName)

        If Not Directory.Exists(FilePath) Then
            Directory.CreateDirectory(FilePath + "\" + FileDate)
        End If

        Select Case FileType
            Case "txt"
                File.AppendAllText(FullPath + "." + FileType, FileData.Trim + vbNewLine)
            Case Else
                File.WriteAllBytes(FullPath, Convert.FromBase64String(FileData.Trim))
        End Select
    End Sub

    Public Function ReadLog(ByVal FilePath As String, ByVal FileName As String) As String
        On Error Resume Next
        ReadLog = ""
        Dim FileDate As String = Now.ToString("yyyy-MM-dd")

        If File.Exists(FilePath + FileName & ".txt") Then
            Using reader As New StreamReader(FilePath + FileName & ".txt", False)
                ReadLog = reader.ReadToEnd
            End Using
        End If
        Return ReadLog
    End Function

    Public Sub LogAllRowsFromDataTable(ByVal dataTable As DataTable, ByVal FilePath As String, ByVal FileName As String)
        On Error Resume Next
        Dim FileDate As String = Now.ToString("yyyy-MM-dd")
        Dim logEntry As New StringBuilder()

        For Each dataRow As DataRow In dataTable.Rows
            For Each col As DataColumn In dataTable.Columns
                logEntry.AppendFormat("{0}: {1}, ", col.ColumnName, If(IsDBNull(dataRow(col)), "NULL", dataRow(col).ToString()))
            Next
            If logEntry.Length > 2 Then logEntry.Length -= 2

            If Not Directory.Exists(FilePath) Then
                Directory.CreateDirectory(FilePath + "\" + FileDate)
            End If

            SaveLog(FilePath, FileName, logEntry.ToString(), "txt")
        Next
    End Sub

    Public Sub ExportTextFromDataTable(ByVal dataTable As DataTable, ByVal FilePath As String, ByVal FileName As String)
        On Error Resume Next
        Dim stringBuilder = New StringBuilder()
        stringBuilder.AppendLine(String.Join(vbTab, dataTable.Columns.Cast(Of DataColumn)().[Select](Function(arg) arg.ColumnName)))

        For Each dataRow As DataRow In dataTable.Rows
            stringBuilder.AppendLine(String.Join(vbTab, dataRow.ItemArray))
        Next

        If Directory.Exists(FilePath) = False Then
            Directory.CreateDirectory(FilePath)
        End If
        File.WriteAllText(FilePath + FileName, stringBuilder.ToString())
    End Sub

    Public Function ReadExcelData(filePath As String) As List(Of String)
        Dim items As New List(Of String)()
        Try
            Dim fileExtension As String = Path.GetExtension(filePath)
            Dim connString As String = If(fileExtension = ".xls",
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filePath & ";Extended Properties='Excel 8.0;HDR=Yes;'",
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties='Excel 12.0 Xml;HDR=Yes;'")

            Using conn As New OleDbConnection(connString)
                conn.Open()
                Dim dtSheets As DataTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                If dtSheets Is Nothing OrElse dtSheets.Rows.Count = 0 Then
                    Return items
                End If

                Dim sheetName As String = dtSheets.Rows(0)("TABLE_NAME").ToString()

                Dim cmd As New OleDbCommand("SELECT EBS FROM [" & sheetName & "]", conn)
                Dim adapter As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable()
                adapter.Fill(dt)

                For Each row As DataRow In dt.Rows
                    items.Add(row(0))
                    'items.Add(String.Join(" | ", row.ItemArray))
                Next
            End Using
        Catch ex As Exception
            MsgBox("Error: ReadExcelData " & ex.Message, MsgBoxStyle.Critical)
            items.Add("Error: " & ex.Message)
        End Try
        Return items
    End Function

    Public Function ReadExcelDataReturnTable(filePath As String) As DataTable
        Dim items As New List(Of String)()
        Try
            Dim fileExtension As String = Path.GetExtension(filePath)
            Dim connString As String = If(fileExtension = ".xls",
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filePath & ";Extended Properties='Excel 8.0;HDR=Yes;'",
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties='Excel 12.0 Xml;HDR=Yes;'")

            Using conn As New OleDbConnection(connString)
                conn.Open()
                Dim dtSheets As DataTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                If dtSheets Is Nothing OrElse dtSheets.Rows.Count = 0 Then
                    Return Nothing
                End If

                Dim sheetName As String = dtSheets.Rows(0)("TABLE_NAME").ToString()

                Dim cmd As New OleDbCommand("SELECT * FROM [" & sheetName & "]", conn)
                Dim adapter As New OleDbDataAdapter(cmd)
                Dim dt As New DataTable()
                adapter.Fill(dt)
                Return dt
            End Using
        Catch ex As Exception
            MsgBox("Error: ReadExcelDataReturnTable " & ex.Message, MsgBoxStyle.Critical)
            Return Nothing
        End Try
    End Function
    'Public Function ImageToByteArray(ByVal imageIn As System.Drawing.Image) As Byte()
    '    Using ms As New MemoryStream()
    '        imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
    '        Return ms.ToArray()
    '    End Using
    'End Function
    Public Function GetMacAddress() As String
        Try
            Dim adapters As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()
            Dim adapter As NetworkInterface
            Dim myMac As String = String.Empty

            For Each adapter In adapters
                ' Check if the adapter is a durable interface (LAN or WLAN)
                If adapter.NetworkInterfaceType = NetworkInterfaceType.Ethernet OrElse
               adapter.NetworkInterfaceType = NetworkInterfaceType.Wireless80211 Then
                    myMac = adapter.GetPhysicalAddress().ToString()
                    Exit For
                End If
            Next
            Return myMac
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    Public Function GetLocalIPAddress() As String
        Dim hostName As String = Dns.GetHostName()
        Dim ipEntry As IPHostEntry = Dns.GetHostEntry(hostName)
        For Each ipAddress As IPAddress In ipEntry.AddressList
            If ipAddress.AddressFamily = AddressFamily.InterNetwork Then
                Return ipAddress.ToString()
            End If
        Next
        Return ""
    End Function

    Public Async Function SendEmail(ByVal Subject As String, ByVal ToEmail As String, ByVal CCEmail As String, ByVal BCCEmail As String, ByVal EmailData As String, ByVal Attach As Object) As Task(Of String)
        Dim EmailFrom As String = "alert@gulahmed.com"
        Dim EmailPassword As String = "Fh#54321"

        Dim smtpClient = New SmtpClient("mail.gulahmed.com") With {
        .Port = 25,
        .Credentials = New NetworkCredential("gulahmed\alert", EmailPassword),
        .EnableSsl = False
    }

        smtpClient.ServicePoint.ConnectionLeaseTimeout = 0
        smtpClient.ServicePoint.MaxIdleTime = 0

        Dim mailMessage = New MailMessage With {
        .From = New MailAddress(EmailFrom),
        .Subject = Subject,
        .Body = EmailData.Trim,
        .IsBodyHtml = True
    }
        'mailMessage.[To].Add("khurram.jawaid@gulahmed.com")

        mailMessage.[To].Add(ToEmail)

        If Not String.IsNullOrWhiteSpace(CCEmail) Then
            mailMessage.[CC].Add(CCEmail)
        End If

        If Not String.IsNullOrWhiteSpace(BCCEmail) Then
            mailMessage.[Bcc].Add(BCCEmail)
        End If

        If Attach IsNot Nothing Then
            mailMessage.Attachments.Add(New Attachment(Attach))
        End If

        Try
            Await smtpClient.SendMailAsync(mailMessage)
            Return "Email sent successfully!"

        Catch ex As SmtpException
            Return $"Error: {ex.Message}, StatusCode: {ex.StatusCode}, InnerException: {ex.InnerException?.Message}"

        Catch ex As Exception
            Return $"Error: {ex.Message}, InnerException: {ex.InnerException?.Message}"
        End Try

    End Function

    Public Function GenerateTemporaryPassword() As String
        Dim chars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
        Dim random As New Random()
        Dim result As New StringBuilder()
        For i As Integer = 1 To 8 ' Generate an 8-character password
            result.Append(chars(random.Next(chars.Length)))
        Next
        Return result.ToString()
    End Function

    Private Function EmailBody(recipientName As String, requestDetails As String, serverUrl As String, yourName As String) As String
        Dim body As String = String.Format(
      "<html>" & vbCrLf &
      "  <body>" & vbCrLf &
      "    <p>Hi {0},</p>" & vbCrLf &
      "    <p>This email requires your approval for {1}.</p>" & vbCrLf &
      "    <p>Please choose your action:</p>" & vbCrLf &
      "    <div style=""display:flex; flex-direction:column>""" & vbCrLf &
      "    <a href=""{2}Y"" style=""background-color: green; color: white; padding: 10px 15px; border: none; border-radius: 5px;"">Approve</button></a>" & vbCrLf &
      "    <a href=""{2}N"" style=""background-color: red; color: white; padding: 10px 15px; border: none; border-radius: 5px;"">Reject</button></a>" & vbCrLf &
      "    </div>" & vbCrLf &
      "    <p>Thanks</p>" & vbCrLf &
      "    <p>{3}</p>" & vbCrLf &
      "  </body>" & vbCrLf &
      "</html>", recipientName, requestDetails, serverUrl, yourName)
        Return body

        '"    <table style=""border: none;"">" & vbCrLf &
        '"      <tr>" & vbCrLf &
        '"        <td><a href=""{2}Y"" style=""background-color: green; color: white; padding: 10px 15px; border: none; border-radius: 5px;"">Approve</button></a></td>" & vbCrLf &
        '"        <td><a href=""{2}N"" style=""background-color: red; color: white; padding: 10px 15px; border: none; border-radius: 5px;"">Reject</button></a></td>" & vbCrLf &
        '"      </tr>" & vbCrLf &
        '"    </table>" & vbCrLf &
    End Function

    Public Function EmailFormat(ByVal dtSource As DataTable, ByVal EmailMsg As String) As String
        Dim tableRows As New StringBuilder()

        If dtSource IsNot Nothing AndAlso dtSource.Rows.Count > 0 Then
            ' Generate table rows
            For Each row As DataRow In dtSource.Rows
                tableRows.Append("<tr>")
                For Each column As DataColumn In dtSource.Columns
                    tableRows.AppendFormat("<td>{0}</td>", row(column))
                Next
                tableRows.Append("</tr>")
            Next

            ' Generate dynamic table headers from column names
            Dim tableHeaders As New StringBuilder()
            For Each column As DataColumn In dtSource.Columns
                tableHeaders.AppendFormat("<th>{0}</th>", column.ColumnName.Replace("_", " "))
            Next

            Return String.Format("
        <html>
            <head>
                <style>
                    table, th, td {{ border: 1px solid black; font-family: 'Calibri'; font-size: 11px; }} 
                    th, td {{ text-align: center; padding: 5px; }} 
                    table {{ width: 100%; border-collapse: collapse; }}
                </style>
            </head>
            <body>
                <p>Dear Admin,<br><br>
                {2}<br>
                Kindly review.<br><br></p>
                <table>
                    <thead><tr>{0}</tr></thead>
                    <tbody>{1}</tbody>
                </table>
                <p>Regards,<br>Attendance Management System<br>Gul Ahmed Textile Mills</p>
            </body>
        </html>", tableHeaders.ToString(), tableRows.ToString(), EmailMsg)
        Else
            Return String.Format("
        <html>
            <head>
                <style>
                    table, th, td {{ border: 1px solid black; font-family: 'Calibri'; font-size: 11px; }} 
                    th, td {{ text-align: center; padding: 5px; }} 
                    table {{ width: 100%; border-collapse: collapse; }}
                </style>
            </head>
            <body>
                <p>Dear Admin,<br><br>
                {0}<br><br></p>
                <p>Regards,<br>Attendance Management System<br>Gul Ahmed Textile Mills</p>
            </body>
        </html>", EmailMsg)
        End If
    End Function

    Public Shared Function isEmail(inputEmail As String) As Boolean
        Dim strRegex As String = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}" + "\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\" + ".)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
        Dim re As New Regex(strRegex)
        If re.IsMatch(inputEmail) Then
            Return (True)
        Else
            Return (False)
        End If
    End Function

    Public Function ConvertCsvDataToDataTable(csvData As String, ByVal UnitName As String) As DataTable
        ' Remove trailing comma if exists
        If csvData.EndsWith(",") Then
            csvData = csvData.TrimEnd(","c)
        End If

        ' Create the DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("Value", GetType(String))
        dt.Columns.Add("UnitName", GetType(String))

        ' Split values and insert into the DataTable
        Dim values() As String = csvData.Split(","c)
        For Each val As String In values
            If Not String.IsNullOrWhiteSpace(val) Then
                dt.Rows.Add(val, UnitName)
            End If
        Next
        Return dt
    End Function

    Public Function ConvertCsvToDataTable(csv As String) As DataTable
        Dim dt As New DataTable()
        Dim lines() As String = csv.Split(New String() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries)

        If lines.Length > 0 Then
            ' Add columns
            Dim headers() As String = lines(0).Split(","c)
            For Each header As String In headers
                dt.Columns.Add(header.Trim())
            Next

            ' Add rows
            For i As Integer = 1 To lines.Length - 1
                Dim values() As String = lines(i).Split(","c)
                dt.Rows.Add(values)
            Next
        End If

        Return dt
    End Function

    Public Sub LogUnhandledError(formName As String)
        If Err.Number <> 0 Then
            Dim errorMessage As String = Err.Description
            Dim stackTrace As String = ""

            Dim ex = Err.GetException()
            If ex IsNot Nothing Then
                stackTrace = ex.StackTrace
            End If

            ' Log error details to a file
            Dim FilePath As String = "C:\TIS-Logs\"
            Dim FileDate As String = Now.ToString("yyyy-MM-dd")
            Dim FullPath As String = Path.Combine(FilePath, FileDate & " ErrorLogs.txt")
            Dim Spacer As String = vbNewLine + "-------------------" + vbNewLine

            If Not Directory.Exists(FilePath) Then
                Directory.CreateDirectory(FilePath)
            End If

            Dim logEntry As String = Spacer &
            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " - Error " &
            Spacer & Err.Number & ": " &
            Spacer & errorMessage & vbCrLf & stackTrace &
            Spacer & " Form Name : " & formName

            My.Computer.FileSystem.WriteAllText(FullPath, logEntry & vbCrLf, True)

            Err.Clear()
        End If
    End Sub

    Public Sub ExportDatatableToCSVFile(Data As DataTable, FileName As String, FilePath As String)
        On Error Resume Next

        Dim FileDate As String = Now.ToString("yyyy-MM-dd")
        Dim FullPath As String = Path.Combine(FilePath, FileDate & FileName & ".Csv")

        If Not Directory.Exists(FullPath) Then
            Directory.CreateDirectory(FullPath)
        End If

        Using writer As New StreamWriter(FullPath, False, System.Text.Encoding.UTF8)
            ' Write headers
            For col As Integer = 0 To Data.Columns.Count - 1
                writer.Write("""" & Data.Columns(col).ColumnName.Replace("""", """""") & """")
                If col < Data.Columns.Count - 1 Then writer.Write(",")
            Next
            writer.WriteLine()

            ' Write rows
            For Each row As DataRow In Data.Rows
                For col As Integer = 0 To Data.Columns.Count - 1
                    Dim field As String = row(col).ToString()
                    ' Escape quotes by doubling them
                    field = field.Replace("""", """""")
                    ' Wrap every field in quotes
                    writer.Write("""" & field & """")
                    If col < Data.Columns.Count - 1 Then writer.Write(",")
                Next
                writer.WriteLine()
            Next
        End Using
        LogUnhandledError("Cls Usefull Class ExportDatatableToCSVFile")
    End Sub
End Class

