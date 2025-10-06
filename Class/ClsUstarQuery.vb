
Public Class ClsUstarQuery

    Private ObjDbConnection As ClsDbConnection = New ClsDbConnection()
    Private ObjUseFulFunctions As ClsUseFulFunctions = New ClsUseFulFunctions()

    Public Function GetRawAttendanceDataQuery(ByVal AttendanceDate As String, ByVal EBS As String) As String
        'Case When Recognition_Time = ATT_Time then Convert(varchar,Cast(ATT_Time as datetime)) Else  Convert(varchar,CAst(ATT_TimeOut as datetime)) End As Recognition_Time,  
        Return String.Format("SELECT Name, Department, Person_ID, Card_Number, 'Device version is too low' AS Body_Temp, 'Device version is too low' AS Temp_Status, Device_Name, Device_SN, Device_Group, 
        Convert(varchar,FORMAT(CAST(ATT_Time AS datetime), 'dd-MM-yyyy HH:m:ss')) AS ATT_Time, Convert(varchar,FORMAT(CAST(ATT_TimeOut AS datetime), 'dd-MM-yyyy HH:m:ss')) AS ATT_TimeOut, Convert(varchar,FORMAT(CAST(Recognition_Time AS datetime), 'dd-MM-yyyy HH:m:ss')) AS Recognition_Time, 'Face' AS Recognition_Mode,
        'Success' AS Identify_Result,'Living' AS Liveness,    'Within Access Time' AS Access_Time,'Within Valid Date' AS Expiry_Date,'' AS Scene_Photo_URL, '' AS Scene_Photo,
        Case When FORMAT(Convert(datetime,Recognition_Time), 'yyyy-MM-dd HH:mm') = FORMAT(Convert(datetime,ATT_Time), 'yyyy-MM-dd HH:mm') then 'IN' Else 'OUT' End AS Remarks
        FROM(
        SELECT rr.emp_no AS Person_ID, e.name AS Name, e.card_no AS Card_Number, dpt.dep_name AS Department, CONVERT(varchar, DATEADD(MILLISECOND, CAST(RIGHT(rr.recognition_time, 3) AS bigint) - 
        DATEDIFF(MILLISECOND, GETDATE(), GETUTCDATE()), DATEADD(SECOND, CAST(LEFT(rr.recognition_time, 10) AS bigint), '1970-01-01 00:00:00')), 20) AS Recognition_Time, Cast(bga.Timein as varchar(30)) ATT_Time,Cast(bga.TimeOut as varchar(30)) ATT_TimeOut, dv.device_ip AS Device_IP, dv.device_name AS Device_Name, dv.device_key AS Device_SN, dvg.group_name AS Device_Group
        FROM 
        employee AS e 
        RIGHT OUTER JOIN emp_dep_relation AS edr 
        INNER JOIN department AS dpt ON edr.dep_id = dpt.id ON e.id = edr.emp_id 
        RIGHT OUTER JOIN device_group AS dvg 
        INNER JOIN recognition_record AS rr 
        INNER JOIN device AS dv ON CAST(rr.device_key AS varchar) = CAST(dv.device_key AS varchar) ON dvg.id = dv.device_group_id ON e.emp_no = rr.emp_no
        INNER JOIN btnGetRawAttendance bga on e.emp_no = bga.EBS
        WHERE 
            ISNUMERIC(rr.emp_no) = 1 
            AND rr.stat_date <> '20000101') AS a
        WHERE 
        CONVERT(DATE, Recognition_Time) = Cast('{0}' as Date) AND Person_ID in ('{1}') Order by CONVERT(datetime, Recognition_Time)", AttendanceDate, EBS)
    End Function
    'Public Function GetRawAttendanceDataQuery(ByVal AttendanceDate As String, ByVal EBS As String, ByVal Gtm8 As String) As String
    '    Return String.Format("SELECT Person_ID,Name,Card_Number,Department,Recognition_Time,Device_IP,Device_Name,Device_SN,Device_Group,'Device version is too low' AS Body_Temp,'Device version is too low' AS Temp_Status,'' AS ATT_Time,'Face' AS Recognition_Mode,'Success' AS Identify_Result,'Living' AS Liveness,'Within Access Time' AS Access_Time,'Within Valid Date' AS Expiry_Date,'' AS Scene_Photo,'' AS Scene_Photo_URL,
    '    CASE WHEN Device_IP IN ('10.8.15.210','10.8.15.152','10.8.15.218','10.8.15.45','10.8.14.188','10.8.14.214','10.8.15.203','10.8.15.214','10.8.15.217','10.8.15.213','10.8.14.202','10.8.14.215') AND CAST(Recognition_Time AS TIME) BETWEEN '13:00' AND '22:00'  THEN 'Out'
    '            WHEN Device_IP IN ('10.8.15.210','10.8.15.152','10.8.15.218','10.8.15.45','10.8.14.188','10.8.14.214','10.8.15.203','10.8.15.214','10.8.15.217','10.8.15.213','10.8.14.202','10.8.14.215') AND CAST(Recognition_Time AS TIME) BETWEEN '06:30' AND '07:30'  THEN 'In' ELSE CASE WHEN CHARINDEX('-IN-', Device_Name) > 0 THEN 'IN' WHEN CHARINDEX('-Out', Device_Name) > 0 THEN 'OUT' Else 'Both' End  END AS Remarks
    '    FROM (SELECT rr.emp_no AS Person_ID, e.name AS Name, e.card_no AS Card_Number, dpt.dep_name AS Department, CONVERT(VARCHAR, DATEADD(MILLISECOND, CAST(RIGHT(rr.recognition_time, 3) AS BIGINT) - DATEDIFF(MILLISECOND, GETDATE(), GETUTCDATE()), DATEADD(SECOND, CAST(LEFT(rr.recognition_time, 10) AS BIGINT), '1970-01-01 00:00:00')), 20) AS Recognition_Time, dv.device_ip AS Device_IP, 
    '            dv.device_name AS Device_Name, dv.device_key AS Device_SN, dvg.group_name AS Device_Group 
    '    FROM employee AS e RIGHT OUTER JOIN emp_dep_relation AS edr INNER JOIN department AS dpt ON edr.dep_id = dpt.id ON e.id = edr.emp_id 
    '            RIGHT OUTER JOIN device_group AS dvg INNER JOIN recognition_record AS rr INNER JOIN device AS dv ON CAST(rr.device_key AS VARCHAR) = CAST(dv.device_key AS VARCHAR) ON dvg.id = dv.device_group_id 
    '            ON e.emp_no = rr.emp_no
    '        WHERE ISNUMERIC(rr.emp_no) = 1 AND rr.stat_date <> '20000101') AS a
    '    WHERE 
    '    CONVERT(DATE, Recognition_Time)= Cast('{0}' as Date) AND Person_ID in ({1}) ", AttendanceDate, EBS)
    'End Function
    Public Function GetRawAttendanceDataQueryView(ByVal EBSCode As String, ByVal DateFrom As Date, ByVal DateTo As Date) As String
        Return String.Format("SELECT * FROM View_Khurram Where EBS in (Select Value From fn_split_string_to_column('{0}',','))
            AND (CAST(Recognition_Time_IN AS date) between '{1}' and '{2}' or CAST(Recognition_Time_Out AS date) between '{1}' and '{2}')
            order by recognition_time", EBSCode, DateFrom, DateTo)
    End Function

    Public Function GetRawAttendanceFromUstar(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal UnitName As String) As String
        'And rr.emp_no IN (SELECT Value FROM fn_split_string_to_column('" + EBSCode + "', ','))
        Return String.Format("WITH recognition_data AS (
                        SELECT rr.emp_no AS EBS, rr.recognition_time, rr.create_time,d.device_name,d.device_ip,                            
                            CONVERT(varchar, DATEADD(MILLISECOND,CAST(RIGHT(rr.recognition_time, 3) AS bigint) - DATEDIFF(MILLISECOND, GETDATE(), GETUTCDATE()), DATEADD(SECOND, CAST(LEFT(rr.recognition_time, 10) AS bigint), '1970-01-01 00:00:00')), 20) AS Recognition_DateTime
                        FROM 
                            dbo.emp_dep_relation AS edr 
                            INNER JOIN dbo.employee AS e ON edr.emp_id = e.id 
                            INNER JOIN dbo.recognition_record AS rr ON e.emp_no = rr.emp_no 
                            INNER JOIN dbo.device AS d ON CAST(rr.device_key AS varchar) = CAST(d.device_key AS varchar)
                        WHERE 
                            ISNUMERIC(rr.emp_no) = 1 
                            AND rr.stat_date <> '20000101'
                            AND rr.emp_no IN (Select EBSCode From TBL_TIS_EBSCode Where UnitName='{2}')
                            AND d.device_ip not in (Select DevIP From tbl_Exclude_Dev))
                    SELECT Null Employee_ID,EBS,Recognition_DateTime,case when device_name like '%-In%' then 1 when device_name like '%-Out%' then 2  else 0 end TransactionINOut,device_ip,device_name,create_time FROM recognition_data
                    WHERE CAST(Recognition_DateTime AS datetime) BETWEEN '{0}' AND '{1}'
                    ORDER BY recognition_time", DateFrom, DateTo, UnitName)
    End Function

    Public Function GetRawAttendanceFromUstarLive(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal UnitName As String) As String
        Return String.Format("WITH recognition_data AS (
        SELECT rr.emp_no AS EBS,rr.recognition_time,CONVERT(VARCHAR, DATEADD(MILLISECOND, CAST(RIGHT(rr.recognition_time, 3) AS BIGINT) - DATEDIFF(MILLISECOND, GETDATE(), GETUTCDATE()), DATEADD(SECOND, CAST(LEFT(rr.recognition_time, 10) AS BIGINT), '1970-01-01 00:00:00')), 20) AS Recognition_DateTime
        FROM dbo.emp_dep_relation AS edr
        INNER JOIN dbo.employee AS e ON edr.emp_id = e.id
        INNER JOIN dbo.recognition_record AS rr ON e.emp_no = rr.emp_no
        INNER JOIN dbo.device AS d ON CAST(rr.device_key AS VARCHAR) = CAST(d.device_key AS VARCHAR)
        WHERE 
        ISNUMERIC(rr.emp_no) = 1 
        AND rr.stat_date <> '20000101'
        AND rr.emp_no IN (Select EBSCode From TBL_TIS_EBSCode Where UnitName='{2}')
        AND d.device_ip NOT IN (SELECT DevIP FROM tbl_Exclude_Dev))
        SELECT EBS,MAX(CAST(Recognition_DateTime AS DATETIME)) AS Max_Recognition_DateTime,MAX(CAST(Recognition_DateTime AS TIME)) AS Max_Recognition_Time
        FROM recognition_data
        WHERE CAST(Recognition_DateTime AS DATETIME) BETWEEN '{0}' AND '{1}'
        GROUP BY EBS
        ORDER BY Max_Recognition_DateTime;", DateFrom, DateTo, UnitName)
    End Function
    Public Function GetRawAttendanceFromZKT(ByVal DateFrom As Date, ByVal DateTo As Date, ByVal UnitName As String) As String
        Return String.Format("SELECT DISTINCT Null Employee_ID, CAST(ac.pin AS int) AS EBS, ac.event_time AS Recognition_DateTime, CASE WHEN isnull(d .host_status, 0) = 1 THEN 1 ELSE 2 END as TransactionINOut, ac.dev_alias AS device_ip, CASE WHEN isnull(d .host_status, 0) = 1 THEN 'Device-IN' ELSE 'Device-OUT' END AS device_name, create_time
                    FROM acc_transaction AS ac INNER JOIN acc_door AS d ON ac.event_point_name = d.name
                    WHERE (CAST(ac.event_time AS date) between '{0}' and '{1}')
                    AND (isnumeric(ac.pin) = 1) AND (ac.pin in (Select EBSCode From TBL_TIS_EBSCode Where UnitName='{2}'))
                    ORDER BY Recognition_DateTime", DateFrom, DateTo, UnitName)

    End Function

    Public Function GetEmployeeAttendance(ByVal EBSCode As String, ByVal DateFrom As Date, ByVal DateTo As Date) As String
        Return String.Format("SELECT 
                            Ea.EmployeeAttendance_Employee_Id_Fk AS EmpID,
                            E.Employee_Code, 
                            Ea.EmployeeAttendance_DateTimeIn AS DateTimeIn, 
                            Ea.EmployeeAttendance_DateTimeOut AS DateTimeOut, 
                            COALESCE(
                                CASE Ea.EmployeeAttendance_Manual
                                    WHEN 0 THEN 'No' 
                                    WHEN 3 THEN 'Yes' 
                                    WHEN 1 THEN 'Admin App' 
                                    WHEN 2 THEN 'HOD App'
                                    ELSE NULL
                                END, CAST(Ea.EmployeeAttendance_Manual AS VARCHAR)) AS Manuals,
                            Ea.EmployeeAttendance_InActive AS InActive, 
                            Ea.EmployeeAttendance_InActiveDate AS InActiveDate, 
                            U.Users_Name AS CreatedBy, 
                            Ea.EmployeeAttendance_CreatedDate AS CreatedDate, 
                            Ea.IpAddress, 
                            Ea.Remarks
                                FROM    
                            EmployeeAttendance AS Ea INNER JOIN Employee AS E 
                            ON Ea.EmployeeAttendance_Employee_Id_Fk = E.Employee_ID LEFT OUTER JOIN Users U on U.Users_Id=Ea.EmployeeAttendance_CreatedBy
                                WHERE  
                            (CAST(Ea.EmployeeAttendance_DateTimeOut AS DATE) between '{1}' and '{2}' 
                            or
                            CAST(Ea.EmployeeAttendance_DateTimeIn AS DATE) between '{1}' and '{2}')
                            AND E.Employee_Code in (Select Value From fn_split_string_to_column('{0}',','))
                            ORDER BY 
                            Ea.EmployeeAttendance_CreatedDate", EBSCode, DateFrom, DateTo)
    End Function

    Public Function GetAttendanceData(ByVal EBSCode As String, ByVal DateFrom As Date, ByVal DateTo As Date) As String
        Return String.Format("SELECT Ad.AttendanceData_Employee_Id_Fk AS Emp_ID,E.Employee_Code AS EBS, E.Employee_Name,  E.Employee_LeftDate,E.PAYROLL_NAME Payroll, Ad.AttendanceData_OverTime AS ActOT,Ad.AttendanceData_ComplianceOT AS CmplOT,
            CASE 
                WHEN DATEDIFF(MINUTE, CAST(AttendanceData_TimeIn AS datetime), CAST(AttendanceData_TimeOut AS datetime)) < 0 
                THEN DATEDIFF(MINUTE, CAST(AttendanceData_TimeIn AS datetime), CAST(AttendanceData_TimeOut AS datetime)) +1440  
                ELSE DATEDIFF(MINUTE, CAST(AttendanceData_TimeIn AS datetime), CAST(AttendanceData_TimeOut AS datetime))
            END AS WorkHours,
	        Ad.AttendanceData_EntryType AS EType,
            Ad.AttendanceData_ShiftBreakTime AS ShBrk,
            FORMAT(CAST(Ad.AttendanceData_DateFor as date),'dddd dd MMMM yyyy') AS DateFor,
            CAST(Ad.AttendanceData_DutyTimeIN AS time(0)) AS DutyTimeIN,
            CAST(Ad.AttendanceData_TimeIn AS datetime) AS TimeIn,
            Ad.AttendanceData_LateTime AS Late,
            CAST(Ad.AttendanceData_DutyTimeOut AS time(0)) AS DutyTimeOut,
            CAST(Ad.AttendanceData_TimeOut AS datetime) AS TimeOut,
            Ad.AttendanceData_EarlyLeaves AS Early,
            Ad.AttendanceData_ScheduledWorkingMinutes AS SchWorking,
	        round(Ad.AttendanceData_TotalWorkingExcludingShiftbreak,-1) TWESB,round(ad.AttendanceData_DutyWorking,-1) DtyWrkg,
            Ad.AttendanceData_EBSFlag AS EBS, Ad.IsAttendanceLock Lock,Ad.AttendanceData_Manual Man,
            Ad.AttendanceData_SickLeave AS SL,ad.AttendanceData_SPL SPL,Ad.AttendanceData_AnuualLeave AL, Ad.AttendanceData_CasualLeave CL, AttendanceData_CPL_Add CPLAdd,AttendanceData_CPL CPL 
        FROM AttendanceData AS Ad
        INNER JOIN Employee AS E ON E.Employee_ID = Ad.AttendanceData_Employee_Id_Fk
        WHERE 1=1
            AND E.Employee_Code in ('{0}')
	        AND CAST(Ad.AttendanceData_DateFor AS date) between '{1}' and '{2}'
        order by CAST(Ad.AttendanceData_DateFor AS date)", EBSCode, DateFrom, DateTo)
    End Function

    Public Function GetAttendanceData(ByVal DateFor As Date, ByVal Cmpl As Boolean) As String
        If Cmpl = True Then
            Return String.Format("SELECT E.Employee_Code AS EBS, CAST(Ad.AttendanceData_TimeIn1 AS datetime) AS TimeIn,CAST(Ad.AttendanceData_TimeOut1 AS datetime) AS TimeOut    
            FROM AttendanceData AS Ad INNER JOIN Employee AS E ON E.Employee_ID = Ad.AttendanceData_Employee_Id_Fk
            WHERE CAST(Ad.AttendanceData_DateFor AS date)= Cast('{0}' as date) and (cast (AttendanceData_TimeIn1 as time)>'00:00' and cast (AttendanceData_TimeOut1 as time)>'00:00')", DateFor)
        Else
            Return String.Format("SELECT E.Employee_Code AS EBS, CAST(Ad.AttendanceData_TimeIn AS datetime) AS TimeIn,CAST(Ad.AttendanceData_TimeOut AS datetime) AS TimeOut    
            FROM AttendanceData AS Ad INNER JOIN Employee AS E ON E.Employee_ID = Ad.AttendanceData_Employee_Id_Fk
            WHERE CAST(Ad.AttendanceData_DateFor AS date)= Cast('{0}' as date) and (cast (AttendanceData_TimeIn as time)>'00:00' and cast (AttendanceData_TimeOut as time)>'00:00')", DateFor)
        End If
    End Function


    Public Function GetRawAttendanceDataFromUstar(ByVal EbsCode As String, ByVal UnitName As String, ByVal DateFrom As Date, ByVal DateTo As Date, Optional ByVal Live As Boolean = False) As DataTable
        Dim DeleteQry As String = String.Format("Delete From TBL_TIS_EBSCode Where UnitName='{0}'", UnitName)
        ObjDbConnection.ExecuteNonQuery(DeleteQry, ObjDbConnection.ConnectDb(ObjDbConnection.DbUface))
        ObjDbConnection.ExecuteNonQuery(DeleteQry, ObjDbConnection.ConnectDb(UnitName))

        Dim dt As DataTable = New DataTable()
        dt = ObjUseFulFunctions.ConvertCsvDataToDataTable(EbsCode, UnitName)
        ObjDbConnection.BulkInsertIntoSQL(dt, "TBL_TIS_EBSCode", ObjDbConnection.ConnectDb(ObjDbConnection.DbUface))
        ObjDbConnection.BulkInsertIntoSQL(dt, "TBL_TIS_EBSCode", ObjDbConnection.ConnectDb(UnitName))

        dt = Nothing
        If Live = False Then
            dt = ObjDbConnection.ExecuteQueryReturnTable(GetRawAttendanceFromUstar(DateFrom, DateTo, UnitName), ObjDbConnection.ConnectDb(ObjDbConnection.DbUface))
        Else
            dt = ObjDbConnection.ExecuteQueryReturnTable(GetRawAttendanceFromUstarLive(DateFrom, DateTo, UnitName), ObjDbConnection.ConnectDb(ObjDbConnection.DbUface))
        End If
        Return dt

    End Function

    Public Function GetRawAttendanceDataFromZKT(ByVal EbsCode As String, ByVal UnitName As String, ByVal DateFrom As Date, ByVal DateTo As Date) As DataTable
        Dim DeleteQry As String = String.Format("Delete From TBL_TIS_EBSCode Where UnitName='{0}'", UnitName)
        ObjDbConnection.ExecuteNonQuery(DeleteQry, ObjDbConnection.ConnectDb(ObjDbConnection.DbZKT))
        ObjDbConnection.ExecuteNonQuery(DeleteQry, ObjDbConnection.ConnectDb(UnitName))

        Dim dt As DataTable = New DataTable()
        dt = ObjUseFulFunctions.ConvertCsvDataToDataTable(EbsCode, UnitName.Trim)
        ObjDbConnection.BulkInsertIntoSQL(dt, "TBL_TIS_EBSCode", ObjDbConnection.ConnectDb(ObjDbConnection.DbZKT))
        ObjDbConnection.BulkInsertIntoSQL(dt, "TBL_TIS_EBSCode", ObjDbConnection.ConnectDb(UnitName))

        dt = Nothing
        dt = ObjDbConnection.ExecuteQueryReturnTable(GetRawAttendanceFromZKT(DateFrom, DateTo, UnitName), ObjDbConnection.ConnectDb(ObjDbConnection.DbZKT))
        Return dt

    End Function
End Class
