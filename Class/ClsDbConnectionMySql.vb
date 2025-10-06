Imports MySql.Data.MySqlClient
Public Class ClsDbConnectionMySql

    Public ConMySQL As MySqlConnection = New MySqlConnection("Server=10.2.0.24;Database=gtm_erp;User ID=gtm_mis;Password=Scada@mis2024;SslMode=Preferred;")
    Private ObjUseFulFunctions As ClsUseFulFunctions = New ClsUseFulFunctions()

    Public Async Function ExecuteMySqlQueryReturnTable(ByVal Query As String, ByVal DbConn As MySqlConnection) As Task(Of DataTable)
        Dim dt As New DataTable()
        Try
            Using MySqlCmd As MySqlCommand = DbConn.CreateCommand()
                MySqlCmd.CommandText = Query
                Await DbConn.OpenAsync()

                Using reader As MySqlDataReader = Await MySqlCmd.ExecuteReaderAsync()
                    dt.Load(reader)
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Error: ExecuteMySqlQueryReturnTable " & ex.Message, MsgBoxStyle.Critical)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteMySqlQueryReturnTable {ex.Message}")
        Finally
            If DbConn.State = ConnectionState.Open Then
                DbConn.Close()
            End If
        End Try
        Return dt
    End Function

    Public Sub ExecuteMySqlNonQuery(ByVal Query As String, ByVal DbConn As MySqlConnection)
        Dim OpenedHere As Boolean = False
        Try
            If DbConn.State <> ConnectionState.Open Then
                DbConn.Open()
                OpenedHere = True
            End If

            Using MySqlCmd As MySqlCommand = DbConn.CreateCommand()
                MySqlCmd.CommandText = Query
                MySqlCmd.ExecuteNonQuery()
            End Using

        Catch ex As Exception
            MsgBox("Error in ExecuteMySqlNonQuery: " & ex.Message & vbCrLf & "Query: " & Query, MsgBoxStyle.Critical)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteMySqlNonQuery {ex.Message}")
        Finally
            If OpenedHere AndAlso DbConn.State = ConnectionState.Open Then
                DbConn.Close()
            End If
        End Try
    End Sub

    Public Sub ExecuteMySqlNonQuery(ByVal Query As String, ByVal DbConn As MySqlConnection, ByVal Param() As MySqlParameter)
        Dim OpenedHere As Boolean = False
        Try
            If DbConn.State <> ConnectionState.Open Then
                DbConn.Open()
                OpenedHere = True
            End If

            Using MySqlCmd As MySqlCommand = DbConn.CreateCommand()
                MySqlCmd.Parameters.Clear()
                If Param IsNot Nothing Then
                    For Each p As MySqlParameter In Param
                        If p IsNot Nothing Then
                            MySqlCmd.Parameters.Add(p)
                        Else
                            Throw New ArgumentNullException("Oracle Parameter is Null!")
                        End If
                    Next
                Else
                    Throw New ArgumentNullException("Parameter Array is Null!")
                End If
                MySqlCmd.CommandText = Query
                MySqlCmd.ExecuteNonQuery()
            End Using

        Catch ex As Exception
            MsgBox("Error in ExecuteMySqlNonQuery: " & ex.Message & vbCrLf & "Query: " & Query, MsgBoxStyle.Critical)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteMySqlNonQuery {ex.Message}")
        Finally
            If OpenedHere AndAlso DbConn.State = ConnectionState.Open Then
                DbConn.Close()
            End If
        End Try
    End Sub
End Class
