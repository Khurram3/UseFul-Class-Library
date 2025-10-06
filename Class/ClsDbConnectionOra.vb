Imports System.Data.OracleClient
'Imports Oracle.ManagedDataAccess.Client

Public Class ClsDbConnectionOra
    Public ConDbUCSHR As OracleConnection = New OracleConnection("Data Source=(DESCRIPTION= (ADDRESS = (PROTOCOL = TCP)(HOST =172.16.15.20) (PORT = 1521))(CONNECT_DATA = (SERVER = dedicated)(SERVICE_NAME = UCSPRD)));User ID=appshr;Password=appshr")
    Public ConDbINTG As OracleConnection = New OracleConnection("Data Source=(DESCRIPTION= (ADDRESS = (PROTOCOL = TCP)(HOST =192.168.24.50) (PORT = 1521))(CONNECT_DATA = (SERVER = dedicated)(SERVICE_NAME = LIVEINTG)));User ID=mohsin;Password=abidi;")
    Public ConDbPRD As OracleConnection = New OracleConnection("Data Source=(DESCRIPTION= (ADDRESS = (PROTOCOL = TCP)(HOST =192.168.24.11) (PORT = 1521))(CONNECT_DATA = (SERVER = dedicated)(SERVICE_NAME = DBPRD)));User ID=apps;Password=Ap#icO412;")
    Public ConDbUCSApps As OracleConnection = New OracleConnection("Data Source=(DESCRIPTION= (ADDRESS = (PROTOCOL = TCP)(HOST =172.16.15.20) (PORT = 1521))(CONNECT_DATA = (SERVER = dedicated)(SERVICE_NAME = UCSPRD)));User ID=apps;Password=gul_oracle_apps;")
    Public ConDbSWS As OracleConnection = New OracleConnection("Data Source=(DESCRIPTION= (ADDRESS = (PROTOCOL = TCP)(HOST =192.168.24.11) (PORT = 1521))(CONNECT_DATA = (SERVER = dedicated)(SERVICE_NAME = DBPRD)));User ID=SWS_SCALES;Password=SWS1234;")
    Public ConDbUCSAppsx As OracleConnection = New OracleConnection("Data Source=(DESCRIPTION= (ADDRESS = (PROTOCOL = TCP)(HOST =localhost) (PORT = 1521))(CONNECT_DATA = (SERVER = dedicated)(SERVICE_NAME = XE)));User ID=SYS;Password=sys;")
    Public ConDbLocalOra As OracleConnection = New OracleConnection("Data Source=localhost:1521/XE;User ID=DBPRD;Password=dbprd;")
    Public ConDbAPEX_APP As OracleConnection = New OracleConnection("Data Source=172.16.15.44:1521/PDBAPEX;User ID=APEX_APP;Password=Apex321;")

    Private ObjUseFulFunctions As ClsUseFulFunctions = New ClsUseFulFunctions
    Public Async Function ExecuteOraQueryReturnTableAsync(ByVal Query As String, ByVal DbConn As OracleConnection) As Task(Of DataTable)
        Dim dt As New DataTable()
        Try
            Using OraCmd As OracleCommand = DbConn.CreateCommand()
                OraCmd.CommandText = Query
                Await DbConn.OpenAsync()

                Using reader As OracleDataReader = Await OraCmd.ExecuteReaderAsync()
                    dt.Load(reader)
                End Using
            End Using

        Catch ex As Exception
            MsgBox("Error: ExecuteOraQueryReturnTableAsync " & ex.Message, MsgBoxStyle.Critical)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteOraQueryReturnTable {ex.Message}")
        Finally
            If DbConn.State = ConnectionState.Open Then
                DbConn.Close()
            End If
        End Try
        Return dt
    End Function

    Public Function ExecuteOraQueryReturnTable(ByVal Query As String, ByVal DbConn As OracleConnection) As DataTable
        Dim dt As New DataTable()
        Try
            Using OraCmd As OracleCommand = DbConn.CreateCommand()
                OraCmd.CommandText = Query
                DbConn.Open()

                Using reader As OracleDataReader = OraCmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Error: ExecuteOraQueryReturnTable " & ex.Message, MsgBoxStyle.Critical)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteOraQueryReturnTable {ex.Message}")
        Finally
            If DbConn.State = ConnectionState.Open Then
                DbConn.Close()
            End If
        End Try
        Return dt
    End Function

    Public Sub ExecuteOraNonQuery(ByVal Query As String, ByVal DbConn As OracleConnection)
        Try
            If DbConn.State <> ConnectionState.Open Then
                DbConn.Open()
            End If

            Using OraCmd As New OracleCommand(Query, DbConn)
                OraCmd.ExecuteNonQuery()
            End Using

        Catch ex As Exception
            MsgBox($"Error in ExecuteOraNonQuery: {ex.Message}{vbCrLf}Query: {Query}", MsgBoxStyle.Critical)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteOraNonQuery {ex.Message}")
        Finally
            If DbConn.State = ConnectionState.Open Then
                Try
                    DbConn.Close()
                Catch
                    ' Ignore close errors
                End Try
            End If
        End Try
    End Sub

    'Public Function BulInsertOra(ByVal dtBulk As DataTable, ByVal TableName As String, ByVal DbConn As OracleConnection) As Boolean
    '    Try
    '        If DbConn.State <> ConnectionState.Open Then
    '            DbConn.Open()
    '        End If

    '        Using bulkCopy As New OracleBulkCopy(DbConn)
    '            bulkCopy.DestinationTableName = TableName

    '            ' Automatically map columns if names match
    '            For Each column As DataColumn In dtBulk.Columns
    '                bulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName)
    '            Next

    '            bulkCopy.WriteToServer(dtBulk)
    '        End Using

    '        Return True
    '    Catch ex As Exception
    '        MsgBox("Bulk Insert Error: " & ex.Message, MsgBoxStyle.Critical)
    '        Return False
    '    Finally
    '        If DbConn.State = ConnectionState.Open Then
    '            DbConn.Close()
    '        End If
    '    End Try
    'End Function


    Public Function ExecuteOraQueryReturnTableWithParamAi(ByVal query As String, ByVal dbConn As OracleConnection, ByVal paramNames As String(), ByVal paramValues As String()) As DataTable
        Dim dt As New DataTable()

        ' Validate inputs
        If paramNames Is Nothing OrElse paramValues Is Nothing OrElse paramNames.Length <> paramValues.Length Then
            Throw New ArgumentException("Parameter names and values must not be null and should have the same length.")
        End If

        Try
            Using dbConn
                Using cmd As New OracleCommand(query, dbConn)
                    cmd.CommandType = CommandType.Text
                    cmd.CommandTimeout = 30 ' Can be made configurable

                    For i As Integer = 0 To paramNames.Length - 1
                        Dim param As New OracleParameter(paramNames(i), paramValues(i))
                        cmd.Parameters.Add(param)
                    Next

                    If dbConn.State <> ConnectionState.Open Then dbConn.Open()
                    Using adp As New OracleDataAdapter(cmd)
                        adp.Fill(dt)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MsgBox($"Error in ExecuteOraQueryReturnTableWithParamAi: {ex.Message}{vbCrLf} Query: {query}", MsgBoxStyle.Critical)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteOraQueryReturnTableWithParamAi {ex.Message}")
            Throw
        End Try

        Return dt
    End Function

    Public Sub ExecuteOraNonQueryWithParam(ByVal Query As String, ByVal dbConn As OracleConnection, ByVal paramNames As String(), ByVal paramValues As String())
        Dim Connection As OracleConnection = New OracleConnection(dbConn.ConnectionString)
        Try
            Connection.Open()
            Using cmd As New OracleCommand()
                For i As Integer = 0 To paramNames.Length - 1
                    cmd.Parameters.Add(New OracleParameter(paramNames(i), paramValues(i)))
                Next
                cmd.CommandText = Query
                cmd.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            MsgBox("Error: ExecuteOraNonQueryWithParam " & ex.Message, MsgBoxStyle.Critical)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteOraNonQueryWithParam {ex.Message}")
        Finally
            If Connection.State = ConnectionState.Open Then
                Connection.Close()
            End If
        End Try
    End Sub
    Public Sub ExecuteOraNonQueryWithParam(ByVal Query As String, ByVal dbConn As OracleConnection, ByVal Param() As OracleParameter)
        Using Connection As New OracleConnection(dbConn.ConnectionString)
            Try
                Connection.Open()
                Using cmd As New OracleCommand(Query, Connection)
                    cmd.Parameters.Clear()
                    If Param IsNot Nothing Then
                        For Each p As OracleParameter In Param
                            If p IsNot Nothing Then
                                cmd.Parameters.Add(p)
                            Else
                                Throw New ArgumentNullException("Oracle Parameter is Null!")
                            End If
                        Next
                    Else
                        Throw New ArgumentNullException("Parameter Array is Null!")
                    End If

                    cmd.ExecuteNonQuery()
                End Using
            Catch ex As Exception
                MsgBox("Error: ExecuteOraNonQueryWithParam " & ex.Message, MsgBoxStyle.Critical)
                ObjUseFulFunctions.LogUnhandledError($"ExecuteOraNonQueryWithParam {ex.Message}")
            Finally
                If Connection.State = ConnectionState.Open Then
                    Connection.Close()
                End If
            End Try
        End Using
    End Sub

    Public Function ExecuteOraRefCursor(ByVal ProcName As String, ByVal dbConn As OracleConnection, ByVal Param() As OracleParameter) As DataTable
        Dim dt As New DataTable()

        Using Connection As New OracleConnection(dbConn.ConnectionString)
            Try
                Connection.Open()
                Using cmd As New OracleCommand(ProcName, Connection)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.Clear()

                    If Param IsNot Nothing Then
                        For Each p As OracleParameter In Param
                            If p IsNot Nothing Then
                                cmd.Parameters.Add(p)
                            Else
                                Throw New ArgumentNullException("Oracle Parameter is Null!")
                            End If
                        Next
                    End If

                    ' Add the OUT ref cursor parameter explicitly (old OracleClient uses Cursor type)
                    Dim outParam As New OracleParameter("p_result", OracleType.Cursor)
                    outParam.Direction = ParameterDirection.Output
                    cmd.Parameters.Add(outParam)

                    ' Fill DataTable from cursor
                    Using da As New OracleDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using
            Catch ex As Exception
                MsgBox("Error: ExecuteOraRefCursor " & ex.Message, MsgBoxStyle.Critical)
                ObjUseFulFunctions.LogUnhandledError($"ExecuteOraRefCursor {ex.Message}")
            Finally
                If Connection.State = ConnectionState.Open Then Connection.Close()
            End Try
        End Using

        Return dt
    End Function

    Public Function ConvertDateFormat(ByVal inputDate As String) As String
        Dim parsedDate As Date
        If Date.TryParse(inputDate, parsedDate) Then
            Return parsedDate.ToString("dd-MMM-yyyy", Globalization.CultureInfo.InvariantCulture)
        Else
            Return "01-Jan-1990"
        End If
    End Function

    Public Function GetEmpImg(ByVal EBSCode As String) As String
        Return String.Format("SELECT resize_img(pi.image) AS EmpImg FROM per_images pi, GTM_EMP_PORTAL_V cv WHERE pi.parent_id = cv.person_id AND pi.table_name = 'PER_PEOPLE_F' AND cv.employee_number='{0}'", EBSCode)
    End Function

End Class
