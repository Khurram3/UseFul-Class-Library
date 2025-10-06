Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
'Imports CrystalDecisions.Shared

Namespace Nam_Data
    Public Class DBBridge
        Public Shared Function DBConnection() As String
            Try
                Dim ConDbLocal As SqlConnection = New SqlConnection("Password=gtmtis@2370;Persist Security Info=True;User ID=sa;Initial Catalog=ERPMS;Data Source=(local)")
                Return ConDbLocal.ToString

                'Dim ObjGlobalEncryption As New ClsDbConnection
                'Dim CU As Microsoft.Win32.RegistryKey += Registry.CurrentUser.CreateSubKey("Smart ERP\Connected\" + Application.StartupPath.ToString)
                'With CU
                '    .OpenSubKey("Smart ERP\Connected", True)
                '    ConnectionName = .GetValue("ConnectionName")
                'End With

                'Dim key As String
                'Dim CU1 As Microsoft.Win32.RegistryKey = Registry.CurrentUser.CreateSubKey("Smart ERP")
                'With CU1
                '    .OpenSubKey("Smart ERP", True)
                '    key = .GetValue(ConnectionName)
                'End With

                'Dim CnS As String
                'If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Smart ERP\" + ConnectionName, key, Nothing) Is Nothing Then
                '    With CU
                '        .OpenSubKey(ConnectionName)
                '        CnS = .GetValue(ConnectionName, key)
                '    End With
                'End If

                'Dim Cn As System.Data.Common.DbConnectionStringBuilder = New System.Data.Common.DbConnectionStringBuilder()
                'Cn.ConnectionString = ObjGlobalEncryption.Decrypt(CnS)
                'Return Cn.ConnectionString
            Catch ce As Exception
                Return MsgBox("Unable to get DB Connection string from Connection Seetings. Contact Administrator")
            End Try
        End Function
        'Public Function RptDBConnection() As ConnectionInfo
        '    Try
        '        Dim myConnectionInfo As ConnectionInfo = New ConnectionInfo()
        '        Dim ObjGlobalEncryption As New ClsDbConnection
        '        Dim ConnectionName As String
        '        Dim CU As Microsoft.Win32.RegistryKey = Registry.CurrentUser.CreateSubKey("Smart ERP\Connected\" & ConnectionName + "\" + Application.StartupPath.ToString)
        '        With CU
        '            .OpenSubKey("Smart ERP\Connected", True)
        '            ConnectionName = .GetValue("ConnectionName")
        '        End With

        '        Dim key As String
        '        Dim CU1 As Microsoft.Win32.RegistryKey = Registry.CurrentUser.CreateSubKey("Smart ERP")
        '        With CU1
        '            .OpenSubKey("Smart ERP", True)
        '            key = .GetValue(ConnectionName)
        '        End With

        '        Dim CnS As String
        '        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Smart ERP\" + ConnectionName, key, Nothing) Is Nothing Then
        '            With CU
        '                .OpenSubKey(ConnectionName)
        '                CnS = .GetValue(ConnectionName, key)
        '            End With
        '        End If

        '        Dim Cn As System.Data.Common.DbConnectionStringBuilder = New System.Data.Common.DbConnectionStringBuilder()
        '        Cn.ConnectionString = ObjGlobalEncryption.Decrypt(CnS)
        '        myConnectionInfo.ServerName = Cn("Data Source")
        '        myConnectionInfo.DatabaseName = Cn("Initial Catalog")
        '        myConnectionInfo.UserID = Cn("uid")
        '        myConnectionInfo.Password = Cn("pwd")
        '        Return myConnectionInfo

        '    Catch ce As Exception
        '        Throw New ApplicationException("Unable to get Report DB Connection Information from Config File. Contact Administrator")
        '    End Try
        'End Function
        ''''------------------------------------------------------------------------------------------------------------
        'Public Function RptDBConnection() As ConnectionInfo
        '    Try
        '        Dim myConnectionInfo As ConnectionInfo = New ConnectionInfo()
        '        myConnectionInfo.ServerName = ("192.168.0.250")
        '        myConnectionInfo.DatabaseName = ("UsedClothing")
        '        myConnectionInfo.UserID = ("Khurram")
        '        myConnectionInfo.Password = ("i")
        '        Return myConnectionInfo

        '    Catch ce As Exception
        '        Throw New ApplicationException("Unable to get Report DB Connection Information from Config File. Contact Administrator")
        '    End Try
        'End Function
        'Public Shared Function DBConnection() As String
        '    Try
        '        DBConnection = "Data Source=192.168.0.250;Initial Catalog=UsedClothing;User ID=Khurram;Password=i"
        '        Return DBConnection
        '    Catch ce As Exception
        '        Return DBConnection
        '    End Try
        'End Function
        Public Shared Function isEmail(inputEmail As String) As Boolean
            Dim strRegex As String = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}" + "\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\" + ".)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
            Dim re As New Regex(strRegex)
            If re.IsMatch(inputEmail) Then
                Return (True)
            Else
                Return (False)
            End If
        End Function
        Public Function ExecuteNonQuery(ByVal storedProcedure As String, ByVal param() As SqlParameter) As Integer
            Try
                Return SqlHelper.ExecuteNonQuery(DBConnection(), CommandType.StoredProcedure, storedProcedure, param)
            Catch sq As SqlException
                Throw sq
            End Try
        End Function
        Public Function ExecuteNonQuerywithTrans(ByVal storedProcedure As String, ByVal param() As SqlParameter) As Integer
            Dim conTrans As New SqlConnection(DBConnection())
            Dim sqlTrans As SqlTransaction
            conTrans.Open()
            Dim returnResult As Integer = 0
            sqlTrans = conTrans.BeginTransaction()
            Using sqlTrans
                Try
                    returnResult = SqlHelper.ExecuteNonQuery(sqlTrans, CommandType.StoredProcedure, storedProcedure, param)
                    sqlTrans.Commit()
                    Return returnResult
                Catch sq As Exception
                    sqlTrans.Rollback()
                    Throw sq
                Finally
                    conTrans.Close()
                End Try
            End Using
        End Function
        Public Function ExecuteNonQuery(ByVal storedProcedure As String) As Integer
            Try
                Return SqlHelper.ExecuteNonQuery(DBConnection(), CommandType.StoredProcedure, storedProcedure)
            Catch sq As SqlException
                Throw sq
            End Try
        End Function
        Public Sub ExecutePayrollQuery(ByVal storedProcedure As String, ByVal param() As SqlParameter)
            Try
                SqlHelper.ExecuteNonQuery(DBConnection(), CommandType.StoredProcedure, storedProcedure, param)
            Catch sq As SqlException
                Throw sq
            End Try
        End Sub
        Public Function ExecuteNonQuerySQL(ByVal sqlquery As String) As Integer
            Try
                Return SqlHelper.ExecuteNonQuery(DBConnection(), CommandType.Text, sqlquery)
            Catch sq As SqlException
                Throw sq
            End Try
        End Function
        Public Function ExecuteNonQuerywithTrans(ByVal storedProcedure As String) As Integer
            Dim conTrans As New SqlConnection(DBConnection())
            Dim sqlTrans As SqlTransaction
            conTrans.Open()
            Dim returnResult As Integer = 0
            sqlTrans = conTrans.BeginTransaction()
            Using sqlTrans
                Try
                    returnResult = SqlHelper.ExecuteNonQuery(sqlTrans, CommandType.StoredProcedure, storedProcedure)
                    sqlTrans.Commit()
                    Return returnResult
                Catch sq As Exception
                    sqlTrans.Rollback()
                    Throw sq
                Finally
                    conTrans.Close()
                End Try
            End Using
        End Function
        Public Function ExecuteNonQuerywithTransfromFrontEnd(ByVal sqlTrans As SqlTransaction, ByVal storedProcedure As String, ByVal param() As SqlParameter) As Integer
            Try
                Dim returnResult As Integer = SqlHelper.ExecuteNonQuery(sqlTrans, CommandType.StoredProcedure, storedProcedure, param)
                Return returnResult
            Catch sq As SqlException
                sqlTrans.Rollback()
                Throw sq
            End Try
        End Function
        Public Function ExecuteDataset(ByVal storedProcedure As String, ByVal param() As SqlParameter) As DataSet
            Try
                Return SqlHelper.ExecuteDataset(DBConnection(), CommandType.StoredProcedure, storedProcedure, param)
            Catch sq As SqlException
                Throw sq
            End Try
        End Function
        Public Function CreateCommand(ByVal storedProcedure As String, ByVal ParamArray param() As String) As SqlCommand
            Try
                Dim conection As New SqlConnection(DBConnection())
                Return SqlHelper.CreateCommand(conection, storedProcedure, param)
            Catch sq As SqlException
                Throw sq
            End Try
        End Function
        Public Function ExecuteDatasetSQL(ByVal storedProcedure As String) As DataSet
            Try
                Return SqlHelper.ExecuteDataset(DBConnection(), CommandType.Text, storedProcedure)
            Catch sq As SqlException
                Throw sq
            End Try
        End Function
        Public Function ExecuteDataset(ByVal storedProcedure As String) As DataSet
            Try
                Return SqlHelper.ExecuteDataset(DBConnection(), CommandType.StoredProcedure, storedProcedure)
            Catch sq As SqlException
                Throw sq
            End Try
        End Function
        Public Function ExecuteScalar(ByVal storedProcedure As String, ByVal param() As SqlParameter) As Object
            Try
                Return SqlHelper.ExecuteScalar(DBConnection(), CommandType.StoredProcedure, storedProcedure, param)
            Catch sq As SqlException
                Throw sq
            End Try
        End Function
        Public Function ExecuteScalar(ByVal sqlTrans As SqlTransaction, ByVal storedProcedure As String, ByVal param() As SqlParameter) As Object
            Try
                Return SqlHelper.ExecuteScalar(sqlTrans, CommandType.StoredProcedure, storedProcedure, param)
            Catch sq As SqlException
                Throw sq
            End Try
        End Function
        Public Function ExecuteReader(ByVal storedProcedure As String, ByVal param() As SqlParameter) As SqlDataReader
            Dim reader As SqlDataReader = Nothing
            Try
                reader = SqlHelper.ExecuteReader(DBConnection(), CommandType.StoredProcedure, storedProcedure, param)
                Return reader
            Catch sq As SqlException
                Throw sq
            End Try
        End Function
        Public Function ExecuteReader(ByVal storedProcedure As String) As SqlDataReader
            Dim reader As SqlDataReader = Nothing
            Try
                reader = SqlHelper.ExecuteReader(DBConnection(), CommandType.StoredProcedure, storedProcedure)
                Return reader
            Catch sq As SqlException
                Throw sq
            End Try
        End Function
        Public Function ExecuteReaderSQL(ByVal sqlquery As String) As SqlDataReader
            Dim reader As SqlDataReader = Nothing
            Try
                reader = SqlHelper.ExecuteReader(DBConnection(), CommandType.Text, sqlquery)
                Return reader
            Catch sq As SqlException
                Throw sq
            End Try
        End Function
        Public Function ExecuteNonQuerywithMultipleTrans(ByVal sqlTrans As SqlTransaction, ByVal storedProcedure As String) As Integer
            Dim returnResult As Integer = 0
            Try
                returnResult = SqlHelper.ExecuteNonQuery(sqlTrans, CommandType.StoredProcedure, storedProcedure)
                Return returnResult
            Catch sq As Exception
                Throw sq
            End Try
        End Function
        Public Function ExecuteNonQuerywithMultipleTrans(ByVal sqlTrans As SqlTransaction, ByVal storedProcedure As String, ByVal param() As SqlParameter) As Integer
            Dim returnResult As Integer = 0
            Try
                returnResult = SqlHelper.ExecuteNonQuery(sqlTrans, CommandType.StoredProcedure, storedProcedure, param)
                Return returnResult
            Catch sq As Exception
                Throw sq
            End Try
        End Function

    End Class

End Namespace


