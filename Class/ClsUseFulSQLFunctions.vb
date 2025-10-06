Public Class ClsUseFulSQLFunctions

    Public Function CheckRunningSQLQuries() As String
        Return String.Format("SELECT
        er.session_id AS [Spid],
        ses.login_name AS [Login],
        ses.host_name AS [Host],
	    er.status as [Status],
	    er.command as [Command],
        er.start_time AS [StartTime],GETDATE() CurrentTime,
        CAST(GETDATE() - er.start_time AS TIME) AS [TimeElapsed],
        OBJECT_NAME(st.objectid) AS [ObjectName],
        st.text AS [SQLStatement]
        FROM sys.dm_exec_requests er
        OUTER APPLY sys.dm_exec_sql_text(er.sql_handle) st
        LEFT JOIN sys.dm_exec_sessions ses ON ses.session_id = er.session_id
        WHERE st.text IS NOT NULL;")
    End Function

    Public Function KillAllRunningSqlQuries() As String
        Return String.Format("USE [master];
        DECLARE @kill varchar(8000) = '';  
        SELECT @kill = @kill + 'kill ' + CONVERT(varchar(5), session_id) + ';'  
        FROM sys.dm_exec_sessions
        WHERE database_id  = db_id('ERPMS')
        Print @Kill;
        EXEC(@kill);")
    End Function

    Public Function GetLastExecutedSqlQuries() As String
        Return String.Format("SELECT execquery.last_execution_time AS [Date Time]	,execsql.TEXT AS [Script]
        FROM sys.dm_exec_query_stats AS execquery
        CROSS APPLY sys.dm_exec_sql_text(execquery.sql_handle) AS execsql
        ORDER BY execquery.last_execution_time DESC")
    End Function

    Public Function GetTablesNamesFromDatabase(ByVal DatabaseName As String) As String
        Return String.Format("SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_CATALOG = '{0}' Order by TABLE_NAME;", DatabaseName)
    End Function
    Public Function GetTISUserRights(ByVal UserName As String)
        Return String.Format("SELECT UR.UsersRights_Id, F.Form_Name, U.Users_Name, U.Users_Password, U.Types, U.Users_Compliance Compl,UR.UsersRights_View, UR.UsersRights_ADD, UR.UsersRights_Delete, UR.UsersRights_Edit
                FROM UsersRights AS UR INNER JOIN
                Form AS F ON UR.UsersRights_Form_Id_Fk = F.Form_Id INNER JOIN
                Users AS U ON UR.UsersRights_Users_Id_Fk = U.Users_Id
                Where U.Users_Name like '%{0}%'", UserName)
    End Function
    Public Function GetTISEmployeeDetailFull(ByVal EBSCode As String)
        Dim Query As String = "SELECT E.Employee_Code,E.Employee_Name,E.Employee_AppointmentDate AS AppointmentDate, E.Employee_LeftDate AS LeftDate,E.Employee_InOutDate InOutDate ,ESbm.SalaryBreakupMaster_Salary AS Salary,    
                    D.Department_Name,B.Branch_Name, Dg.Designation_Name, E.PAYROLL_NAME,E.Employee_CreatedDate AS CreatedDate,U.Users_Name AS CreatedBy,EBSA.SALARY_BASIS,E.Employee_SalaryType_Id_Fk AS StypeId,    
                    EBSA.BANK_ACCOUNT_NUMBER, EBSA.ACTUAL_SALARY 
                FROM Employee AS E
                Right outer JOIN EmployeeSalaryBreakup AS ESb ON ESb.EmployeeSalaryBreakup_Employee_Id_Fk = E.Employee_ID
                Right outer JOIN SalaryBreakupMaster AS ESbm ON ESbm.SalaryBreakupMaster_ID = ESb.EmployeeSalaryBreakup_SalaryBreakup_Id_Fk
                Right outer JOIN Employee_EBS_Active AS EBSA ON E.Employee_Code = EBSA.EMPLOYEE_NUMBER
                Right outer JOIN EmployeeSection AS ES ON E.Employee_ID = ES.EmployeeSection_Employee_Id_Fk
                Right outer JOIN EmployeeDesignation AS EDg ON E.Employee_ID = EDg.EmployeeDesignation_Employee_Id_Fk
                Right outer JOIN Designation AS Dg ON Dg.Designation_Id = EDg.EmployeeDesignation_Designation_Id_Fk
                Right outer JOIN EmployeeDepartment AS ED ON E.Employee_ID = ED.EmployeeDepartment_Employee_Id_Fk
                Right outer JOIN Department AS D ON D.Department_Id = ED.EmployeeDepartment_Department_Id_Fk
                Right outer JOIN EmployeeBranch AS EB ON E.Employee_ID = EB.EmployeeBranch_Employee_Id_Fk
                Right outer JOIN Branch AS B ON B.Branch_Id = EB.EmployeeBranch_Branch_Id_Fk
                Right outer JOIN Users AS U ON U.Users_Id = E.Employee_CreatedBy
                WHERE 1=1 "

        If EBSCode = "NULL" Then
            Return String.Format(Query + " AND E.Employee_LeftDate='' OR (YEAR(CAST(E.Employee_LeftDate AS date)) > YEAR(GETDATE()) - 1)")
        End If
        Return String.Format(Query + " AND (e.Employee_Code in (Select Value From fn_split_string_to_column('{0}',',')))", EBSCode)
    End Function
    Public Function GetTISEmployeeDetail(ByVal EBSCode As String)
        Dim Query As String = "SELECT E.Employee_Code,E.Employee_ID,E.Employee_Name,G.Gender_Name Gender,E.Employee_AppointmentDate AS AppointmentDate, E.Employee_LeftDate AS LeftDate,
            D.Department_Name,Dg.Designation_Name, E.PAYROLL_NAME
            FROM Employee AS E
            Right outer JOIN EmployeeDesignation AS EDg ON E.Employee_ID = EDg.EmployeeDesignation_Employee_Id_Fk
            Right outer JOIN Designation AS Dg ON Dg.Designation_Id = EDg.EmployeeDesignation_Designation_Id_Fk
            Right outer JOIN EmployeeDepartment AS ED ON E.Employee_ID = ED.EmployeeDepartment_Employee_Id_Fk
            Right outer JOIN Department AS D ON D.Department_Id = ED.EmployeeDepartment_Department_Id_Fk
            Right outer JOIN Gender As G On E.Employee_Gender_Id_Fk=G.Gender_ID
            WHERE 1=1 "

        If EBSCode = "NULL" Then
            Return String.Format(Query + " AND E.Employee_LeftDate='' OR (YEAR(CAST(E.Employee_LeftDate AS date)) > YEAR(GETDATE()) - 1)")
        End If
        Return String.Format(Query + " AND (e.Employee_Code in (Select Value From fn_split_string_to_column('{0}',',')))", EBSCode)
    End Function
    Public Function CheckDatabaseTableSize()
        Return String.Format("SELECT t.name AS TableName,s.name AS SchemaName,p.rows,
                                SUM(a.total_pages) * 8 AS TotalSpaceKB, 
                                CAST(ROUND(((SUM(a.total_pages) * 8) / 1024.00), 2) AS NUMERIC(36, 2)) AS TotalSpaceMB,
                                CAST(ROUND(((SUM(a.total_pages) * 8) / 1024.00), 2)/1024 AS NUMERIC(36, 2)) AS TotalGB,
                                SUM(a.used_pages) * 8 AS UsedSpaceKB, 
                                CAST(ROUND(((SUM(a.used_pages) * 8) / 1024.00), 2) AS NUMERIC(36, 2)) AS UsedSpaceMB, 
                                CAST(ROUND(((SUM(a.used_pages) * 8) / 1024.00), 2)/1024 AS NUMERIC(36, 2)) AS UsedGB, 
                                (SUM(a.total_pages) - SUM(a.used_pages)) * 8 AS UnusedSpaceKB,
                                CAST(ROUND(((SUM(a.total_pages) - SUM(a.used_pages)) * 8) / 1024.00, 2) AS NUMERIC(36, 2)) AS UnusedSpaceMB
                            FROM sys.tables t
                            INNER JOIN sys.indexes i ON t.object_id = i.object_id
                            INNER JOIN sys.partitions p ON i.object_id = p.object_id AND i.index_id = p.index_id
                            INNER JOIN sys.allocation_units a ON p.partition_id = a.container_id
                            LEFT OUTER JOIN sys.schemas s ON t.schema_id = s.schema_id
                            WHERE t.name NOT LIKE 'dt%' AND t.is_ms_shipped = 0 AND i.object_id > 255 
                            GROUP BY t.name, s.name, p.rows
                            ORDER BY TotalSpaceMB DESC, t.name")
    End Function
    Public Function CheckActiveSQLConnections()
        Return String.Format("Select DB_NAME(dbid) As DBName, COUNT(dbid) As NumberOfConnections,loginame As LoginName FROM sys.sysprocesses WHERE dbid > 0 GROUP BY dbid, loginame;")
    End Function

    Public Function GetMonthly_OT(ByVal DateFrom As String, ByVal DateTo As String, ByVal EbsCode As String, ByVal Departments As String)
        Dim Query As String = "SELECT E.Employee_Code, E.Employee_Name, CAST(E.Employee_AppointmentDate AS date) AS AppointmentDate, E.Employee_LeftDate AS LeftDate, E.Employee_OTEnt, E.PAYROLL_NAME, B.Branch_Name, D.Department_Name, 
                            CE.CategoryEmployee_Name AS [Sub Depart], Dg.Designation_Name, E.PERSON_TYPE, EBS.ACTUAL_SALARY, CAST(Ad.AttendanceData_DateFor AS date) AS DateFor, Shift.Shift_Name, 
                            Ad.AttendanceData_DutyTimeIN AS Shift_TimeIn, Ad.AttendanceData_DutyTimeOut AS Shift_TimeOut, Cast(Ad.AttendanceData_TimeIn as date) DateIn, Cast(Ad.AttendanceData_TimeIn as time(0)) AS TimeIn, Ad.AttendanceData_LateTime LateHrs, CAst(Ad.AttendanceData_TimeOut as date) AS DateOut, Cast(Ad.AttendanceData_TimeOut  as time(0)) TimeOut,
                            Ad.AttendanceData_DutyWorking DutyWork,Ad.AttendanceData_EarlyLeaves ShortHrs,Round(Ad.AttendanceData_ShiftBreakTime / 60, 1) ShiftBreak,Ad.AttendanceData_BreakTime BreakTime, Ad.AttendanceData_TotalWorkingExcludingShiftbreak TWESB, Ad.AttendanceData_AcctualOverTime AS OT_Hrs, Ad.AttendanceData_EntryType AS EType,ad.AttendanceData_CPL_Add CPL_Add 
                            FROM Employee AS E
                            LEFT JOIN EmployeeBranch AS EB ON E.Employee_ID = EB.EmployeeBranch_Employee_Id_Fk
                            LEFT JOIN Branch AS B ON EB.EmployeeBranch_Branch_Id_Fk = B.Branch_Id
                            LEFT JOIN EmployeeDepartment AS ED ON E.Employee_ID = ED.EmployeeDepartment_Employee_Id_Fk
                            LEFT JOIN Department AS D ON ED.EmployeeDepartment_Department_Id_Fk = D.Department_Id
                            LEFT JOIN EmployeeDesignation AS Edg ON E.Employee_ID = Edg.EmployeeDesignation_Employee_Id_Fk
                            LEFT JOIN Designation AS Dg ON Edg.EmployeeDesignation_Designation_Id_Fk = Dg.Designation_Id
                            LEFT JOIN EmployeeCategory AS EC ON E.Employee_ID = EC.EmployeeCategory_Employee_Id_Fk
                            LEFT JOIN CategoryEmployee AS CE ON EC.EmployeeCategory_CategoryEmployee_Id_Fk = CE.CategoryEmployee_ID
                            LEFT JOIN AttendanceData AS Ad ON E.Employee_ID = Ad.AttendanceData_Employee_Id_Fk
                            LEFT JOIN Shift ON Ad.AttendanceData_ShiftID = Shift.Shift_ID
                            LEFT JOIN Employee_EBS_Active AS EBS ON E.Employee_Code = EBS.EMPLOYEE_NUMBER
                            WHERE        (1 = 1) 
                            AND ((E.Employee_LeftDate = '') or (CAST(E.Employee_LeftDate AS Date)>= '{0}'))
                            AND (CAST(Ad.AttendanceData_DateFor AS Date) BETWEEN '{0}' AND '{1}')"


        If EbsCode <> "" Then
            Query += " And (e.Employee_Code in (Select Value From fn_split_string_to_column(isnull('{2}',e.Employee_Code),','))) "
            'MsgBox(EbsCode)
        End If

        If Departments <> "" Then
            Query += " AND (D.Department_Name in (Select Value From fn_split_string_to_column(isnull('{3}',D.Department_Name),','))) "
            'MsgBox(Departments)
        End If
        Return String.Format(Query, DateFrom, DateTo, EbsCode, Departments)
    End Function

    Public Function CheckSQLJobHistory()
        Return String.Format("USE msdb;
                            SELECT j.name AS JobName,h.run_date AS RunDate,h.run_time AS RunTime,h.run_duration AS RunDuration,CASE h.run_status
                                    WHEN 0 THEN 'Failed'
                                    WHEN 1 THEN 'Succeeded'
                                    WHEN 2 THEN 'Retry'
                                    WHEN 3 THEN 'Canceled'
                                    WHEN 4 THEN 'In Progress'
                                    ELSE 'Unknown' END AS RunStatus,h.message AS Message
                            FROM dbo.sysjobs j INNER JOIN dbo.sysjobhistory h ON j.job_id = h.job_id
                            ORDER BY h.run_date DESC, h.run_time DESC;")
    End Function

    Public Function CheckSQLJobsCreated()
        Return String.Format("SELECT j.name AS JobName,s.name AS ScheduleName,s.enabled AS ScheduleEnabled,js.next_run_date,js.next_run_time
                FROM msdb.dbo.sysjobs j
                LEFT JOIN msdb.dbo.sysjobschedules js ON j.job_id = js.job_id
                LEFT JOIN     msdb.dbo.sysschedules s ON js.schedule_id = s.schedule_id
                ORDER BY j.name;")
    End Function
    Public Function AttendanceInOutMissingSummary(ByVal DateFrom As String, ByVal DateTo As String) As String
        Return String.Format("SELECT E.PAYROLL_NAME, E.PERSON_TYPE, Ad.AttendanceData_EntryType AS EType, CAST(Ad.AttendanceData_DateFor AS date) AS DateFor, COUNT(Ad.AttendanceData_EntryType) AS EQty, B.Branch_Name
                            FROM EmployeeBranch AS EB INNER JOIN
                                    Branch AS B ON EB.EmployeeBranch_Branch_Id_Fk = B.Branch_Id INNER JOIN
                                    AttendanceData AS Ad INNER JOIN
                                    Employee AS E ON E.Employee_ID = Ad.AttendanceData_Employee_Id_Fk ON EB.EmployeeBranch_Employee_Id_Fk = E.Employee_ID
                            WHERE        (CAST(Ad.AttendanceData_DateFor AS date) BETWEEN '{0}' AND '{1}')
                            GROUP BY B.Branch_Name,
                            Ad.AttendanceData_EntryType, 
                            E.PAYROLL_NAME, E.PERSON_TYPE, 
                            CAST(Ad.AttendanceData_DateFor AS date)
                            HAVING        ((Ad.AttendanceData_EntryType IN('II', 'IO', 'IHI', 'IHO')))
                            ORDER BY CAST(Ad.AttendanceData_DateFor AS date)", DateFrom, DateTo)
    End Function

    Public Function CheckRunningSyncSerivceQuery(ByVal Status As String, ByVal DateFrom As String, ByVal DateTo As String, Optional ByVal UnitName As String = "", Optional ServiceName As String = "Ebs Sync Service") As String
        Dim Qry As String = ""

        Qry = String.Format("SELECT ApLog.UnitId, Info.Instance, Info.UnitName, ApLog.ProcessName, Cast(ApLog.starttime as varchar(50)) starttime, Cast(ApLog.endtime as Varchar(50)) endtime, DATEDIFF(MINUTE,ApLog.starttime,ApLog.endtime) Duration, ApLog.Status, ApLog.Message
                            FROM ApplicationsLogs AS ApLog INNER JOIN InstanceInfo AS Info ON ApLog.UnitId = Info.id
                            WHERE (Cast(ApLog.starttime as date) Between '{1}' AND '{2}') 
                            AND ApLog.Status=ISNULL('{0}',Status) 
                            ", Status, DateFrom, DateTo)

        If UnitName <> "" Then
            Qry += " And info.UnitName ='" & UnitName & "' "
        End If

        Select Case ServiceName
            Case "Ebs Sync Service"
                Qry += " AND ProcessName = 'Ebs Sync Service' "
            Case "TIS To EBS Attendance Log"
                Qry += " AND ProcessName = 'TIS To EBS Attendance Log' "
            Case ""

        End Select

        Qry += " Order by Cast(ApLog.starttime as date) desc"
        Return Qry
    End Function

    Public Function GetEmailAlerts(ByVal UnitName As String)
        If UnitName <> "" Then
            Return String.Format("SELECT Ea.Id, Ea.AlertName, Ea.EmailSubject, Ea.UnitName, Ea.EmailToRecipent, Ea.EmailCCRecipent,Ea.EmailBCCRecipent, Ea.EmailDetailText , EaHR.HRBPEmail, EaHR.CCEmail FROM aa_alertsdetail AS Ea LEFT OUTER JOIN aa_alertsdetailHRBP AS EaHR ON Ea.UnitName = EaHR.UnitName WHERE (isnull(Ea.InActive,0) = 0) AND (isnull(EaHR.IsActive,0) = 0) And Ea.UnitName='{0}' order by UnitName ", UnitName)
        Else
            Return String.Format("SELECT Ea.Id, Ea.AlertName, Ea.EmailSubject, Ea.UnitName, Ea.EmailToRecipent, Ea.EmailCCRecipent,Ea.EmailBCCRecipent, Ea.EmailDetailText , EaHR.HRBPEmail, EaHR.CCEmail FROM aa_alertsdetail AS Ea LEFT OUTER JOIN aa_alertsdetailHRBP AS EaHR ON Ea.UnitName = EaHR.UnitName WHERE (isnull(Ea.InActive,0) = 0) AND (isnull(EaHR.IsActive,0) = 0) order by UnitName ")
        End If
    End Function

    Public Function GetEmailAlertsByInstance(ByVal InstanceName As String)
        If (InstanceName <> "") Then
            Return String.Format("SELECT Ea.Id, Ea.AlertName, Ea.EmailSubject, Ea.UnitName, Ea.EmailToRecipent, Ea.EmailCCRecipent, Ea.EmailBCCRecipent, Ea.EmailDetailText, EaHR.HRBPEmail, EaHR.CCEmail
            FROM aa_alertsdetail AS Ea 
            LEFT OUTER JOIN aa_alertsdetailHRBP AS EaHR ON Ea.UnitName = EaHR.UnitName
            Inner Join InstanceInfo I on I.UnitName=Ea.UnitName
            WHERE (ISNULL(Ea.InActive, 0) = 0) AND (ISNULL(EaHR.IsActive, 0) = 0) 
            AND AlertName like '%Unlocked%'
            AND (I.Instance = '{0}')
            ORDER BY Ea.UnitName", InstanceName)
        Else
            Return String.Format("SELECT Ea.Id, Ea.AlertName, Ea.EmailSubject, Ea.UnitName, Ea.EmailToRecipent, Ea.EmailCCRecipent, Ea.EmailBCCRecipent, Ea.EmailDetailText, EaHR.HRBPEmail, EaHR.CCEmail
            FROM aa_alertsdetail AS Ea 
            LEFT OUTER JOIN aa_alertsdetailHRBP AS EaHR ON Ea.UnitName = EaHR.UnitName
            Inner Join InstanceInfo I on I.UnitName=Ea.UnitName
            WHERE (ISNULL(Ea.InActive, 0) = 0) AND (ISNULL(EaHR.IsActive, 0) = 0) 
            AND AlertName like '%Unlocked%'
            AND (I.Instance = '{0}')
            ORDER BY Ea.UnitName", InstanceName)
        End If
    End Function

    Public Function GetEmailAlertsTime(ByVal AlertName As String)
        If AlertName <> "" Then
            Return String.Format("SELECT Id,AlertName,DayName, SechduleTime,IsActive FROM aa_alertsdetailTiming Where IsActive=0 And AlertName= '{0}' Order by Id", AlertName)
        Else
            Return String.Format("SELECT Id,AlertName,DayName, SechduleTime,IsActive FROM aa_alertsdetailTiming Where IsActive=0 Order by Id")
        End If
    End Function

    Public Function GetShrinkDatabaseLog() As String
        Return String.Format("SELECT 'USE [' + d.name + N']' + CHAR(13) + CHAR(10) + 'DBCC SHRINKFILE (N''' + mf.name + N''' , 0, TRUNCATEONLY)' + CHAR(13) + CHAR(10) + CHAR(13) + CHAR(10) 
                              FROM sys.master_files mf JOIN sys.databases d ON mf.database_id = d.database_id 
                              WHERE d.database_id > 4;")
    End Function

    Public Function GetBackupDetails(Optional ByVal DbName As String = "") As String
        Return String.Format("WITH LatestBackups AS (
        SELECT 
            a.server_name AS ServerName,
            a.database_name AS DatabaseName,
            Cast(a.backup_finish_date as date) AS BackupDate,
            a.backup_size / 1024 / 1024 AS BackupSizeMb,
            CASE a.[type]
                WHEN 'D' THEN 'Full'
                WHEN 'I' THEN 'Differential'
                WHEN 'L' THEN 'Transaction Log'
                ELSE a.[type]
            END AS BackupType,
            b.physical_device_name AS PhysicalAddress,
            ROW_NUMBER() OVER (
                PARTITION BY a.database_name 
                ORDER BY a.backup_finish_date DESC
            ) AS rn
        FROM 
            msdb.dbo.backupset a
        INNER JOIN 
            msdb.dbo.backupmediafamily b ON a.media_set_id = b.media_set_id
    )
    SELECT ServerName,DatabaseName,BackupDate,BackupSizeMb,BackupType,PhysicalAddress FROM LatestBackups WHERE rn = 1 ORDER BY BackupDate DESC;")
    End Function

    Public Function Checkfragmentation() As String
        Return String.Format("SELECT dbschemas.[name] AS SchemaName, dbtables.[name] AS TableName, dbindexes.[name] AS IndexName, 
        Round(indexstats.avg_fragmentation_in_percent,0) Avg_Fragment
        FROM sys.dm_db_index_physical_stats (DB_ID(), NULL, NULL, NULL, 'LIMITED') indexstats
        INNER JOIN sys.tables dbtables ON dbtables.[object_id] = indexstats.[object_id]
        INNER JOIN sys.schemas dbschemas ON dbtables.[schema_id] = dbschemas.[schema_id]
        INNER JOIN sys.indexes AS dbindexes ON dbindexes.[object_id] = indexstats.[object_id]
            AND indexstats.index_id = dbindexes.index_id
        WHERE indexstats.database_id = DB_ID()
        ORDER BY indexstats.avg_fragmentation_in_percent DESC;")
    End Function

    Public Function CheckDatabaseAndLogFileSize() As String
        Return String.Format("SELECT 
        db.name AS DatabaseName,
        mf.name AS FileName,
        mf.type_desc AS FileType,
        CAST(mf.size * 8.0 / 1024 AS DECIMAL(10,2)) AS FileSizeMB,
        CAST(mf.size * 8.0 / 1048576 AS DECIMAL(10,2)) AS FileSizeGB
        FROM sys.master_files mf
        INNER JOIN sys.databases db ON db.database_id = mf.database_id
        ORDER BY db.name, FileType")
    End Function
End Class
