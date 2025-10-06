Imports System.Data.SqlClient
Imports MiniProjectV1.Nam_Data

Public Class Clsaa_alertsdetail
#Region "Variables"
    Private Const _spName As String = "SP_aa_alertsdetail"
    Private ObjDBBridge As New DBBridge

    Private _Id As Integer
    Private _AlertName As String
    Private _UnitName As String
    Private _EmailToRecipent As String
    Private _EmailCCRecipent As String
    Private _EmailBCCRecipent As String
    Private _OracleQuery As String
    Private _SQLQUERY As String
    Private _SqlInstanceId As Integer
    Private _DbRunType As Integer
    Private _RoutineType As Integer
    Private _RoutineName As String
    Private _IsFromStartDays As Boolean
    Private _IsFromEndDays As Boolean
    Private _StartTime As DateTime
    Private _RepeatMinute As Integer
    Private _InActive As Boolean
    Private _EmailSubject As String
    Private _EmailDetailText As String
    Private _RandomDaysDates As String
    Private _LastGetDatetimeTime As DateTime
    Private _QuryTypeSlct_Ins_Upd_Dlt As Integer

#End Region

#Region "Public Property"

    Public Property Id() As Integer
        Get
            Return _Id
        End Get
        Set(ByVal value As Integer)
            _Id = value
        End Set
    End Property

    Public Property AlertName() As String
        Get
            Return _AlertName
        End Get
        Set(ByVal value As String)
            _AlertName = value
        End Set
    End Property

    Public Property UnitName() As String
        Get
            Return _UnitName
        End Get
        Set(ByVal value As String)
            _UnitName = value
        End Set
    End Property

    Public Property EmailToRecipent() As String
        Get
            Return _EmailToRecipent
        End Get
        Set(ByVal value As String)
            _EmailToRecipent = value
        End Set
    End Property

    Public Property EmailCCRecipent() As String
        Get
            Return _EmailCCRecipent
        End Get
        Set(ByVal value As String)
            _EmailCCRecipent = value
        End Set
    End Property

    Public Property EmailBCCRecipent() As String
        Get
            Return _EmailBCCRecipent
        End Get
        Set(ByVal value As String)
            _EmailBCCRecipent = value
        End Set
    End Property

    Public Property OracleQuery() As String
        Get
            Return _OracleQuery
        End Get
        Set(ByVal value As String)
            _OracleQuery = value
        End Set
    End Property

    Public Property SQLQUERY() As String
        Get
            Return _SQLQUERY
        End Get
        Set(ByVal value As String)
            _SQLQUERY = value
        End Set
    End Property

    Public Property SqlInstanceId() As Integer
        Get
            Return _SqlInstanceId
        End Get
        Set(ByVal value As Integer)
            _SqlInstanceId = value
        End Set
    End Property

    Public Property DbRunType() As Integer
        Get
            Return _DbRunType
        End Get
        Set(ByVal value As Integer)
            _DbRunType = value
        End Set
    End Property

    Public Property RoutineType() As Integer
        Get
            Return _RoutineType
        End Get
        Set(ByVal value As Integer)
            _RoutineType = value
        End Set
    End Property

    Public Property RoutineName() As String
        Get
            Return _RoutineName
        End Get
        Set(ByVal value As String)
            _RoutineName = value
        End Set
    End Property

    Public Property IsFromStartDays() As Boolean
        Get
            Return _IsFromStartDays
        End Get
        Set(ByVal value As Boolean)
            _IsFromStartDays = value
        End Set
    End Property

    Public Property IsFromEndDays() As Boolean
        Get
            Return _IsFromEndDays
        End Get
        Set(ByVal value As Boolean)
            _IsFromEndDays = value
        End Set
    End Property

    Public Property StartTime() As DateTime
        Get
            Return _StartTime
        End Get
        Set(ByVal value As DateTime)
            _StartTime = value
        End Set
    End Property

    Public Property RepeatMinute() As Integer
        Get
            Return _RepeatMinute
        End Get
        Set(ByVal value As Integer)
            _RepeatMinute = value
        End Set
    End Property

    Public Property InActive() As Boolean
        Get
            Return _InActive
        End Get
        Set(ByVal value As Boolean)
            _InActive = value
        End Set
    End Property

    Public Property EmailSubject() As String
        Get
            Return _EmailSubject
        End Get
        Set(ByVal value As String)
            _EmailSubject = value
        End Set
    End Property

    Public Property EmailDetailText() As String
        Get
            Return _EmailDetailText
        End Get
        Set(ByVal value As String)
            _EmailDetailText = value
        End Set
    End Property

    Public Property RandomDaysDates() As String
        Get
            Return _RandomDaysDates
        End Get
        Set(ByVal value As String)
            _RandomDaysDates = value
        End Set
    End Property

    Public Property LastGetDatetimeTime() As DateTime
        Get
            Return _LastGetDatetimeTime
        End Get
        Set(ByVal value As DateTime)
            _LastGetDatetimeTime = value
        End Set
    End Property

    Public Property QuryTypeSlct_Ins_Upd_Dlt() As Integer
        Get
            Return _QuryTypeSlct_Ins_Upd_Dlt
        End Get
        Set(ByVal value As Integer)
            _QuryTypeSlct_Ins_Upd_Dlt = value
        End Set
    End Property


#End Region

#Region "Method"
    Public Function Insert() As Integer
        Dim param(22) As SqlParameter
        param(0) = New SqlParameter("@Mode", "Insert")
        param(1) = New SqlParameter("@Id", _Id)
        param(2) = New SqlParameter("@AlertName", _AlertName)
        param(3) = New SqlParameter("@UnitName", _UnitName)
        param(4) = New SqlParameter("@EmailToRecipent", _EmailToRecipent)
        param(5) = New SqlParameter("@EmailCCRecipent", _EmailCCRecipent)
        param(6) = New SqlParameter("@EmailBCCRecipent", _EmailBCCRecipent)
        param(7) = New SqlParameter("@OracleQuery", _OracleQuery)
        param(8) = New SqlParameter("@SQLQUERY", _SQLQUERY)
        param(9) = New SqlParameter("@SqlInstanceId", _SqlInstanceId)
        param(10) = New SqlParameter("@DbRunType", _DbRunType)
        param(11) = New SqlParameter("@RoutineType", _RoutineType)
        param(12) = New SqlParameter("@RoutineName", _RoutineName)
        param(13) = New SqlParameter("@IsFromStartDays", _IsFromStartDays)
        param(14) = New SqlParameter("@IsFromEndDays", _IsFromEndDays)
        param(15) = New SqlParameter("@StartTime", _StartTime)
        param(16) = New SqlParameter("@RepeatMinute", _RepeatMinute)
        param(17) = New SqlParameter("@InActive", _InActive)
        param(18) = New SqlParameter("@EmailSubject", _EmailSubject)
        param(19) = New SqlParameter("@EmailDetailText", _EmailDetailText)
        param(20) = New SqlParameter("@RandomDaysDates", _RandomDaysDates)
        param(21) = New SqlParameter("@LastGetDatetimeTime", _LastGetDatetimeTime)
        param(22) = New SqlParameter("@QuryTypeSlct_Ins_Upd_Dlt", _QuryTypeSlct_Ins_Upd_Dlt)

        Return ObjDBBridge.ExecuteNonQuery(_spName, param)

    End Function

    Public Function Update() As Integer
        Dim param(22) As SqlParameter
        param(0) = New SqlParameter("@Mode", "Update")
        param(1) = New SqlParameter("@Id", _Id)
        param(2) = New SqlParameter("@AlertName", _AlertName)
        param(3) = New SqlParameter("@UnitName", _UnitName)
        param(4) = New SqlParameter("@EmailToRecipent", _EmailToRecipent)
        param(5) = New SqlParameter("@EmailCCRecipent", _EmailCCRecipent)
        param(6) = New SqlParameter("@EmailBCCRecipent", _EmailBCCRecipent)
        param(7) = New SqlParameter("@OracleQuery", _OracleQuery)
        param(8) = New SqlParameter("@SQLQUERY", _SQLQUERY)
        param(9) = New SqlParameter("@SqlInstanceId", _SqlInstanceId)
        param(10) = New SqlParameter("@DbRunType", _DbRunType)
        param(11) = New SqlParameter("@RoutineType", _RoutineType)
        param(12) = New SqlParameter("@RoutineName", _RoutineName)
        param(13) = New SqlParameter("@IsFromStartDays", _IsFromStartDays)
        param(14) = New SqlParameter("@IsFromEndDays", _IsFromEndDays)
        param(15) = New SqlParameter("@StartTime", _StartTime)
        param(16) = New SqlParameter("@RepeatMinute", _RepeatMinute)
        param(17) = New SqlParameter("@InActive", _InActive)
        param(18) = New SqlParameter("@EmailSubject", _EmailSubject)
        param(19) = New SqlParameter("@EmailDetailText", _EmailDetailText)
        param(20) = New SqlParameter("@RandomDaysDates", _RandomDaysDates)
        param(21) = New SqlParameter("@LastGetDatetimeTime", _LastGetDatetimeTime)
        param(22) = New SqlParameter("@QuryTypeSlct_Ins_Upd_Dlt", _QuryTypeSlct_Ins_Upd_Dlt)

        Return ObjDBBridge.ExecuteNonQuery(_spName, param)

    End Function

    Public Function Delete() As Integer
        Dim param(22) As SqlParameter
        param(0) = New SqlParameter("@Mode", "Delete")
        param(1) = New SqlParameter("@Id", _Id)
        param(2) = New SqlParameter("@AlertName", _AlertName)
        param(3) = New SqlParameter("@UnitName", _UnitName)
        param(4) = New SqlParameter("@EmailToRecipent", _EmailToRecipent)
        param(5) = New SqlParameter("@EmailCCRecipent", _EmailCCRecipent)
        param(6) = New SqlParameter("@EmailBCCRecipent", _EmailBCCRecipent)
        param(7) = New SqlParameter("@OracleQuery", _OracleQuery)
        param(8) = New SqlParameter("@SQLQUERY", _SQLQUERY)
        param(9) = New SqlParameter("@SqlInstanceId", _SqlInstanceId)
        param(10) = New SqlParameter("@DbRunType", _DbRunType)
        param(11) = New SqlParameter("@RoutineType", _RoutineType)
        param(12) = New SqlParameter("@RoutineName", _RoutineName)
        param(13) = New SqlParameter("@IsFromStartDays", _IsFromStartDays)
        param(14) = New SqlParameter("@IsFromEndDays", _IsFromEndDays)
        param(15) = New SqlParameter("@StartTime", _StartTime)
        param(16) = New SqlParameter("@RepeatMinute", _RepeatMinute)
        param(17) = New SqlParameter("@InActive", _InActive)
        param(18) = New SqlParameter("@EmailSubject", _EmailSubject)
        param(19) = New SqlParameter("@EmailDetailText", _EmailDetailText)
        param(20) = New SqlParameter("@RandomDaysDates", _RandomDaysDates)
        param(21) = New SqlParameter("@LastGetDatetimeTime", _LastGetDatetimeTime)
        param(22) = New SqlParameter("@QuryTypeSlct_Ins_Upd_Dlt", _QuryTypeSlct_Ins_Upd_Dlt)

        Return ObjDBBridge.ExecuteNonQuery(_spName, param)

    End Function

    Public Sub SelectById()
        Dim param(1) As SqlParameter
        param(0) = New SqlParameter("@Mode", "ViewById")
        param(1) = New SqlParameter("@Id", _Id)
        Dim dt As New DataTable()
        dt = ObjDBBridge.ExecuteDataset(_spName, param).Tables(0)
        If dt.Rows.Count <> 0 Then
            Dim dr As DataRow
            dr = dt.Rows(0)

            _Id = IIf(IsDBNull(dr("Id")) = True, Nothing, dr("Id"))
            _AlertName = dr("AlertName").ToString
            _UnitName = dr("UnitName").ToString
            _EmailToRecipent = dr("EmailToRecipent").ToString
            _EmailCCRecipent = dr("EmailCCRecipent").ToString
            _EmailBCCRecipent = dr("EmailBCCRecipent").ToString
            _OracleQuery = dr("OracleQuery").ToString
            _SQLQUERY = dr("SQLQUERY").ToString
            _SqlInstanceId = IIf(IsDBNull(dr("SqlInstanceId")) = True, Nothing, dr("SqlInstanceId"))
            _DbRunType = IIf(IsDBNull(dr("DbRunType")) = True, Nothing, dr("DbRunType"))
            _RoutineType = IIf(IsDBNull(dr("RoutineType")) = True, Nothing, dr("RoutineType"))
            _RoutineName = dr("RoutineName").ToString
            _IsFromStartDays = Convert.ToBoolean(dr("IsFromStartDays"))
            _IsFromEndDays = Convert.ToBoolean(dr("IsFromEndDays"))
            _StartTime = Convert.ToDateTime(dr("StartTime"))
            _RepeatMinute = IIf(IsDBNull(dr("RepeatMinute")) = True, Nothing, dr("RepeatMinute"))
            _InActive = Convert.ToBoolean(dr("InActive"))
            _EmailSubject = dr("EmailSubject").ToString
            _EmailDetailText = dr("EmailDetailText").ToString
            _RandomDaysDates = dr("RandomDaysDates").ToString
            _LastGetDatetimeTime = Convert.ToDateTime(dr("LastGetDatetimeTime"))
            _QuryTypeSlct_Ins_Upd_Dlt = IIf(IsDBNull(dr("QuryTypeSlct_Ins_Upd_Dlt")) = True, Nothing, dr("QuryTypeSlct_Ins_Upd_Dlt"))

        End If

    End Sub

    Public Function GetId() As Integer
        Dim param(0) As SqlParameter
        param(0) = New SqlParameter("@Mode", "GetId")
        Dim dt As New DataTable()
        dt = ObjDBBridge.ExecuteDataset(_spName, param).Tables(0)
        If dt.Rows.Count <> 0 Then
            Dim dr As DataRow
            dr = dt.Rows(0)
            GetId = dr(0).ToString
        End If
        Return GetId
    End Function

    Public Function DBPaneGrid() As BindingSource
        Dim param(0) As SqlParameter
        param(0) = New SqlParameter("@Mode", "PaneGrid")
        Dim dt As New DataTable
        dt = ObjDBBridge.ExecuteDataset(_spName, param).Tables(0)
        Dim DataBind As New BindingSource
        DataBind.DataSource = dt
        Return DataBind
    End Function

    Public Function DBLovAlertName() As DataTable
        Dim param(0) As SqlParameter
        param(0) = New SqlParameter("@Mode", "LovAlertName")
        DBLovAlertName = ObjDBBridge.ExecuteDataset(_spName, param).Tables(0)
        Return DBLovAlertName
    End Function

#End Region
End Class
