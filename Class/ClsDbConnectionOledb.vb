Imports System.Data.OleDb
Imports System.IO

Public Class ClsDbConnectionOledb
    Private ObjUseFulFunctions As ClsUseFulFunctions = New ClsUseFulFunctions()

    Public Function ConnectDb(ByVal ExcelFile As String, ByVal Ver As String) As OleDbConnection
        Try
            If Not File.Exists(ExcelFile) Then
                Throw New FileNotFoundException("The specified Excel file does not exist: " & ExcelFile)
            End If
            Dim provider As String
            Select Case Ver
                Case "4"
                    provider = "Microsoft.Jet.OLEDB.4.0"
                Case "12", "15", "16"
                    provider = "Microsoft.ACE.OLEDB.12.0"
                Case Else
                    provider = "Microsoft.ACE.OLEDB.12.0"
            End Select

            Dim connectionString As String = "Provider=" & provider & ";Data Source=" & ExcelFile & ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;'"
            Return New OleDbConnection(connectionString)
        Catch ex As Exception
            MsgBox("Error in ConnectDb OleDbConnection: " & ex.Message, MsgBoxStyle.Critical)
            ObjUseFulFunctions.LogUnhandledError($"ExecuteMySqlNonQuery {ex.Message}")
            Return Nothing
        End Try
    End Function


    Public Sub ExecuteNonQuery(ByVal Query As String, ByVal DbConn As OleDbConnection)
        Dim Connection As OleDbConnection = DbConn
        Try
            Connection.Open()
            Dim cmd As New OleDbCommand
            cmd.CommandText = Query
            cmd.CommandTimeout = 0
            cmd.Connection = Connection
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message + "  ExecuteNonQuery", MsgBoxStyle.Information)
        Finally
            If Connection.State = ConnectionState.Open Then
                Connection.Close()
            End If
        End Try
    End Sub

    Public Function ExecuteQueryReturnTable(ByVal Query As String, ByVal DbConn As OleDbConnection) As DataTable
        Dim ds As DataSet = New DataSet
        Dim da As OleDbDataAdapter = New OleDbDataAdapter(Query, DbConn)
        Try
            DbConn.Open()
            da.SelectCommand.CommandTimeout = 0
            da.Fill(ds, 0)
            ExecuteQueryReturnTable = ds.Tables(0)
        Catch ex As Exception
            MsgBox(ex.Message + "  ExecuteQueryReturnTable", MsgBoxStyle.Information)
            Return Nothing
        Finally
            da.Dispose()
            If DbConn.State = ConnectionState.Open Then
                DbConn.Close()
            End If
            If DbConn.State = ConnectionState.Open Then
                DbConn.Close()
            End If
        End Try
        Return ExecuteQueryReturnTable
    End Function
End Class
