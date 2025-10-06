Imports System.Data.SqlClient
Imports System.Text
Imports System.Security.Cryptography
Imports System.IO

Public Class ClsDbConnection

    Public ConDbSystain As SqlConnection = New SqlConnection("Server=10.2.52.33\SQLEXPRESS;Database=Exchange;User Id=farhan;Password=abc$123;")
    Public ConDbLocal As SqlConnection = New SqlConnection("Data Source=(local);Initial Catalog=ERPMS;User ID=sa;Password=gtmtis@2370;Persist Security Info=True;")

    Public DbVisistor As String = "172.16.2.26\Visitor"
    Public DbUface As String = "172.16.2.29\Uface"
    Public DbZKT As String = "172.16.2.44\Security"
    Public DbCMS As String = "172.16.2.26\CMS"
    Public DbERPMS_Helper As String = "172.16.2.26\General"

    Private ObjUseFulFunctions As ClsUseFulFunctions = New ClsUseFulFunctions()
    Public Function ConnectDb(ByVal Servers As String, Optional ByVal UserAs As String = "TIS") As SqlConnection
        Dim connectionString As String = ""

        Select Case Servers

            Case "172.16.2.29\Uface"
                connectionString = "Data Source=" & Servers & ";Initial Catalog=luna;User ID=sa;Password=Uface123"

            Case "172.16.2.44\Security"
                connectionString = "Data Source=" & Servers & ";Initial Catalog=ZKT;User ID=sa;Password=123123"

            Case "172.16.2.26\CMS"
                If UserAs = "TIS" Then
                    connectionString = "Data Source=" & Servers & ";Initial Catalog=CMSDBNEW;User ID=TIS;Password=GtM@2024$Secure!"
                Else
                    connectionString = "Data Source=" & Servers & ";Initial Catalog=CMSDBNEW;User ID=sa;Password=gtmtis@2370"
                End If

            Case "172.16.2.26\General"
                If UserAs = "TIS" Then
                    connectionString = "Data Source=" & Servers & ";Initial Catalog=ERPMS_Helper;User ID=TIS;Password=GtM@2024$Secure!"
                Else
                    connectionString = "Data Source=" & Servers & ";Initial Catalog=ERPMS_Helper;User ID=sa;Password=gtmtis@2370"
                End If

            Case Else
                If UserAs = "TIS" Then
                    connectionString = "Data Source=" & Servers & ";Initial Catalog=ERPMS;User ID=TIS;Password=GtM@2024$Secure!"
                Else
                    connectionString = "Data Source=" & Servers & ";Initial Catalog=ERPMS;User ID=sa;Password=gtmtis@2370"
                End If

        End Select

        Return New SqlConnection(connectionString)
    End Function

    Public Function ConnectServer(ByVal Servers As String, ByVal DbName As String) As SqlConnection
        Dim connectionString As String = ""
        Select Case Servers.ToUpper()
            Case ("172.16.2.29\Uface").ToUpper()
                connectionString = "Data Source=" & Servers & ";Initial Catalog=luna;User ID=sa;Password=Uface123"
            Case ("172.16.2.44\Security").ToUpper
                connectionString = "Data Source=" & Servers & ";Initial Catalog=ZKT;User ID=sa;Password=123123"
            Case Else
                connectionString = "Data Source=" & Servers & ";Initial Catalog=" & DbName & ";User ID=sa;Password=gtmtis@2370"
        End Select
        Return New SqlConnection(connectionString)
    End Function

    Public Sub ExecuteNonQuery(ByVal Query As String, ByVal DbConn As SqlConnection)
        Dim Connection As SqlConnection = DbConn
        Try
            Connection.Open()
            Dim cmd As New SqlCommand
            cmd.CommandText = Query
            cmd.CommandTimeout = 0
            cmd.Connection = Connection
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message + "  ExecuteNonQuery", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteNonQuery {ex.Message}")
        Finally
            If Connection.State = ConnectionState.Open Then
                Connection.Close()
            End If
        End Try
    End Sub
    Public Sub ExecuteNonQuery(ByVal Query As String, ByVal DbConn As SqlConnection, param() As SqlParameter)
        Dim Connection As SqlConnection = DbConn
        Try
            Connection.Open()
            Dim cmd As New SqlCommand
            cmd.CommandText = Query
            cmd.CommandTimeout = 0
            cmd.Connection = Connection
            If param.Length > 0 Then
                For i = 0 To param.Length - 1 Step 1
                    cmd.Parameters.AddWithValue(param(i).ParameterName, param(i).Value)
                Next
            End If
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message + "  ExecuteNonQuery", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteNonQuery {ex.Message}")
        Finally
            If Connection.State = ConnectionState.Open Then
                Connection.Close()
            End If
        End Try
    End Sub
    Public Sub ExecuteNonQuerySp(ByVal Query As String, ByVal DbConn As SqlConnection, param() As SqlParameter)
        Dim Connection As SqlConnection = DbConn
        Try
            Connection.Open()
            Dim Cmd As SqlCommand
            Cmd = Connection.CreateCommand()
            Cmd.CommandText = Query
            Cmd.CommandTimeout = 0
            Cmd.CommandType = CommandType.StoredProcedure
            If param.Length > 0 Then
                For i = 0 To param.Length - 1 Step 1
                    Cmd.Parameters.AddWithValue(param(i).ParameterName, param(i).Value)
                Next
            End If
            Cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message + "  ExecuteNonQuerySp", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteNonQuerySp {ex.Message}")
        Finally
            If Connection.State = ConnectionState.Open Then
                Connection.Close()
            End If
        End Try
    End Sub
    Public Sub ExecuteNonQuerySp(ByVal Query As String, ByVal DbConn As SqlConnection, ByVal Param As String(), ByVal ParamVal As String())
        Try
            DbConn.Open()
            Dim Cmd As SqlCommand
            Cmd = DbConn.CreateCommand()
            Cmd.CommandText = Query
            Cmd.CommandTimeout = 0
            Cmd.CommandType = CommandType.StoredProcedure
            If Param.Length > 0 Then
                For i = 0 To Param.Length - 1 Step 1
                    Cmd.Parameters.AddWithValue(Param(i), ParamVal(i))
                Next
            End If
            Cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message + "  ExecuteNonQuerySp", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteNonQuerySp {ex.Message}")
        Finally
            If DbConn.State = ConnectionState.Open Then
                DbConn.Close()
            End If
        End Try
    End Sub
    Public Function ExecuteQueryReturnTable(ByVal Query As String, ByVal DbConn As SqlConnection) As DataTable
        Dim ds As DataSet = New DataSet
        Dim da As SqlDataAdapter = New SqlDataAdapter(Query, DbConn)
        Try
            DbConn.Open()
            da.SelectCommand.CommandTimeout = 0
            da.Fill(ds, 0)
            ExecuteQueryReturnTable = ds.Tables(0)
        Catch ex As Exception
            MsgBox(ex.Message + "  ExecuteQueryReturnTable", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteQueryReturnTable {ex.Message}")
            Return ds.Tables(0)
        Finally
            da.Dispose()
            If DbConn.State = ConnectionState.Open Then
                DbConn.Close()
            End If
        End Try
        Return ExecuteQueryReturnTable
    End Function
    Public Function ExecuteQueryReturnTable(ByVal Query As String, ByVal DbConn As SqlConnection, ByVal Param As String(), ByVal ParamVal As String()) As DataTable
        Dim dt As New DataTable()
        Try
            Using cmd As New SqlCommand(Query, DbConn)
                cmd.CommandTimeout = 0
                If Param IsNot Nothing AndAlso Param.Length > 0 Then
                    For i As Integer = 0 To Param.Length - 1
                        cmd.Parameters.AddWithValue(Param(i), ParamVal(i))
                    Next
                End If

                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using

            End Using
        Catch ex As Exception
            MsgBox(ex.Message + "  ExecuteQueryReturnTable", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteQueryReturnTable {ex.Message}")
        Finally
            If DbConn.State = ConnectionState.Open Then
                DbConn.Close()
            End If
        End Try
        Return dt
    End Function

    Public Function ExecuteQueryReturnTableWithParam(ByVal Query As String, ByVal DbConn As SqlConnection, ByVal Param As String(), ByVal ParamVal As String()) As DataTable
        Dim cmd As SqlCommand = New SqlCommand(Query, DbConn)
        Dim ds As DataSet = New DataSet()
        Try
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = 0
            cmd.Connection = DbConn
            For i = 0 To Param.Length - 1 Step 1
                cmd.Parameters.AddWithValue(Param(i), ParamVal(i))
            Next
            DbConn.Open()
            Dim adp As New SqlDataAdapter(cmd)
            cmd.ExecuteNonQuery()
            adp.Fill(ds, "Table")
        Catch ex As Exception
            MsgBox(ex.Message + "  ExecuteQueryReturnTableWithParam", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteQueryReturnTableWithParam {ex.Message}")
            Return ds.Tables(0)
        Finally
            If DbConn.State = ConnectionState.Open Then
                DbConn.Close()
            End If
        End Try
        Return ds.Tables(0)
    End Function
    Public Function ExecuteQueryReturnTableWithParam(ByVal query As String, ByVal dbConn As SqlConnection, ByVal parameters() As SqlParameter) As DataTable
        Dim dt As New DataTable()
        Try
            Using cmd As New SqlCommand(query, dbConn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.CommandTimeout = 0

                If parameters IsNot Nothing Then
                    cmd.Parameters.AddRange(parameters)
                End If

                Using adp As New SqlDataAdapter(cmd)
                    adp.Fill(dt)
                End Using
            End Using
        Catch ex As Exception
            MsgBox(ex.Message + "  ExecuteQueryReturnTableWithParam", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteQueryReturnTableWithParam: {ex.Message}")
        Finally
            If dbConn.State = ConnectionState.Open Then
                dbConn.Close()
            End If
        End Try
        Return dt
    End Function

    Public Function ExecuteQueryReturnTableWithParam1(ByVal Query As String, ByVal DbConn As SqlConnection, ByVal Param() As SqlParameter)
        Dim cmd As SqlCommand = New SqlCommand(Query, DbConn)
        Dim ds As DataSet = New DataSet()
        Try
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = 0
            cmd.Connection = DbConn

            If Param IsNot Nothing Then
                cmd.Parameters.AddRange(Param)
            End If

            DbConn.Open()
            Dim adp As New SqlDataAdapter(cmd)
            cmd.ExecuteNonQuery()
            adp.Fill(ds, "Table")
        Catch ex As Exception
            MsgBox(ex.Message + "  ExecuteQueryReturnTableWithParam1", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteQueryReturnTableWithParam {ex.Message}")
            Return ds.Tables(0)
        Finally
            If DbConn.State = ConnectionState.Open Then
                DbConn.Close()
            End If
        End Try
        Return ds.Tables(0)
    End Function
    Public Function ExecuteQueryReturnTableWithParamAi(ByVal query As String, ByVal dbConn As SqlConnection, ByVal Param() As SqlParameter) As DataTable
        Dim dt As New DataTable()
        Try
            Using cmd As New SqlCommand(query, dbConn)
                cmd.CommandType = CommandType.Text
                cmd.CommandTimeout = 0

                If Param IsNot Nothing Then
                    cmd.Parameters.AddRange(Param)
                End If

                dbConn.Open()
                Using adp As New SqlDataAdapter(cmd)
                    adp.Fill(dt)
                End Using
            End Using
        Catch ex As Exception
            MsgBox(ex.Message + "  ExecuteQueryReturnTableWithParamAi", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteQueryReturnTableWithParamAi {ex.Message}")
        Finally
            If dbConn.State = ConnectionState.Open Then
                dbConn.Close()
            End If
        End Try
        Return dt
    End Function
    Public Function ExecuteQueryReturnTableWithParamAi(ByVal query As String, ByVal dbConn As SqlConnection, ByVal paramNames As String(), ByVal paramValues As String()) As DataTable
        Dim dt As New DataTable()
        Try
            Using cmd As New SqlCommand(query, dbConn)
                cmd.CommandType = CommandType.Text
                cmd.CommandTimeout = 0
                For i As Integer = 0 To paramNames.Length - 1
                    cmd.Parameters.Add(New SqlParameter(paramNames(i), paramValues(i)))
                Next
                dbConn.Open()
                Using adp As New SqlDataAdapter(cmd)
                    adp.Fill(dt)
                End Using
            End Using
        Catch ex As Exception
            MsgBox(ex.Message + "  ExecuteQueryReturnTableWithParamAi", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteQueryReturnTableWithParamAi {ex.Message}")
        Finally
            If dbConn.State = ConnectionState.Open Then
                dbConn.Close()
            End If
        End Try
        Return dt
    End Function
    Public Function ExecuteQuerySpReturnSqlDataReader(ByVal Query As String, ByVal DbConn As SqlConnection, ByVal Param() As SqlParameter) As SqlDataReader
        Try
            If DbConn.State <> ConnectionState.Open Then
                DbConn.Open()
            End If
            Dim cmd As New SqlCommand(Query, DbConn)
            cmd.CommandType = CommandType.StoredProcedure
            For Each parameter As SqlParameter In Param
                cmd.Parameters.AddWithValue(parameter.ParameterName, parameter.Value)
            Next
            Dim reader As SqlDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            Return reader
        Catch ex As Exception
            MsgBox(ex.Message + "  ExecuteQuerySpReturnSqlDataReader", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteQuerySpReturnSqlDataReader {ex.Message}")
            Return Nothing
        End Try
    End Function

    Public Function ExecuteQuerySpReturnScalar(ByVal Query As String, ByVal DbConn As SqlConnection, ByVal Param() As SqlParameter) As Object
        Try
            If DbConn.State <> ConnectionState.Open Then
                DbConn.Open()
            End If

            Using cmd As New SqlCommand(Query, DbConn)
                cmd.CommandType = CommandType.StoredProcedure

                For Each parameter As SqlParameter In Param
                    cmd.Parameters.AddWithValue(parameter.ParameterName, parameter.Value)
                Next

                Dim result As Object = cmd.ExecuteScalar()
                Return result
            End Using

        Catch ex As Exception
            MsgBox(ex.Message + "  ExecuteQuerySpReturnScalar", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteQuerySpReturnScalar {ex.Message}")
            Return Nothing
        Finally
            If DbConn.State = ConnectionState.Open Then
                DbConn.Close()
            End If
        End Try
    End Function

    Public Function ExecuteQueryReturnScalar(ByVal Query As String, ByVal DbConn As SqlConnection) As Object
        Try
            If DbConn.State <> ConnectionState.Open Then
                DbConn.Open()
            End If

            Using cmd As New SqlCommand(Query, DbConn)
                cmd.CommandType = CommandType.Text

                Dim result As Object = cmd.ExecuteScalar()
                Return result
            End Using

        Catch ex As Exception
            MsgBox(ex.Message + "  ExecuteQueryReturnScalar", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteQueryReturnScalar {ex.Message}")
            Return Nothing
        Finally
            If DbConn.State = ConnectionState.Open Then
                DbConn.Close()
            End If
        End Try
    End Function


    Public Function Encrypt(clearText As String) As String
        Dim EncryptionKey As String = "MK"
        Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, &H65, &H64, &H76, &H65, &H64, &H65, &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                    cs.Write(clearBytes, 0, clearBytes.Length)
                    cs.Close()
                End Using
                clearText = Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
        Return clearText
    End Function
    Public Function Decrypt(cipherText As String) As String
        Dim EncryptionKey As String = "MK"
        Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D, &H65, &H64, &H76, &H65, &H64, &H65, &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                    cs.Write(cipherBytes, 0, cipherBytes.Length)
                    cs.Close()
                End Using
                cipherText = Encoding.Unicode.GetString(ms.ToArray())
            End Using
        End Using
        Return cipherText
    End Function
    Public Function CheckConnectDbState(ByVal Server As String, ByVal ConnectionType As String, ByVal UserID As String, ByVal Password As String) As String
        Dim Connection As String
        If ConnectionType = "Windows" Then
            Connection = "Data Source=" + Server + ";Initial Catalog=master;Integrated Security=SSPI"
        Else
            Connection = "Data Source=" + Server + ";Initial Catalog=master" + ";User ID=" + UserID + ";Password=" + Password + ";"
        End If
        Return Connection
    End Function
    Public Function GetDataBaseList(ByVal Server As String, ByVal ConnectionType As String, ByVal UserID As String, ByVal Password As String) As DataTable
        Dim dt As New DataTable()
        Dim Connection As SqlConnection = New SqlConnection
        Try
            Connection.ConnectionString = CheckConnectDbState(Server, ConnectionType, UserID, Password)
            dt = ExecuteQueryReturnTable("SELECT Name FROM master.dbo.sysdatabases WHERE dbid > 3 order by name", Connection, Nothing, Nothing)
            Return dt
        Catch ex As Exception
            MsgBox(ex.Message + "  GetDataBaseList", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"GetDataBaseList {ex.Message}")
            Return Nothing
        End Try
    End Function
    Public Function GetTableList(ByVal Server As String, ByVal ConnectionType As String, ByVal UserID As String, ByVal Password As String) As DataTable
        Dim dt As New DataTable()
        Dim Connection As SqlConnection = New SqlConnection
        Try
            Connection.ConnectionString = CheckConnectDbState(Server, ConnectionType, UserID, Password)
            dt = ExecuteQueryReturnTable("Select Name, Create_Date, Modify_Date from sys.tables", Connection, Nothing, Nothing)
            Return dt
        Catch ex As Exception
            MsgBox(ex.Message + "  GetTableList", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"GetTableList {ex.Message}")
            Return Nothing
        End Try
    End Function
    Public Function GetColoumnList(ByVal Server As String, ByVal ConnectionType As String, ByVal UserID As String, ByVal Password As String, ByVal TableName As String) As DataTable
        Dim dt As New DataTable()
        Dim Connection As SqlConnection = New SqlConnection
        Try
            Connection.ConnectionString = CheckConnectDbState(Server, ConnectionType, UserID, Password)
            dt = ExecuteQueryReturnTable("SELECT column_name ,data_type,character_maximum_length FROM INFORMATION_SCHEMA.COLUMNS WHERE table_name = '" + TableName + "'", Connection, Nothing, Nothing)
            Return dt
        Catch ex As Exception
            MsgBox(ex.Message + "  GetColoumnList", MsgBoxStyle.Information)
            ObjUseFulFunctions.LogUnhandledError($"GetColoumnList {ex.Message}")
            Return Nothing
        End Try
    End Function
    Public Sub BulkInsertIntoSQL(ByVal dataTable As DataTable, ByVal TableName As String, ByVal DbConn As SqlConnection)
        If DbConn.State <> ConnectionState.Open Then
            DbConn.Open()
        End If

        Using bulkCopy As New SqlBulkCopy(DbConn)
            bulkCopy.DestinationTableName = TableName
            bulkCopy.BulkCopyTimeout = 0
            Try
                bulkCopy.WriteToServer(dataTable)
            Catch ex As Exception
                MsgBox("Bulk Insert Error: " & ex.Message, MsgBoxStyle.Information)
                ObjUseFulFunctions.LogUnhandledError($"BulkInsertIntoSQL {ex.Message}")
            Finally
                If DbConn.State = ConnectionState.Open Then
                    DbConn.Close()
                End If
            End Try
        End Using
    End Sub


End Class
