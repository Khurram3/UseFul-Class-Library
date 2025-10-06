Public Class ClsOraQuery
    Public Function GetPayrollQuery(ByVal DateFrom As String, ByVal DateTo As String, IsFnf As Integer) As String
        Dim Query As String = String.Format("SELECT
            cb.empid           AS Ebs,
            cb.attand_date_frm As Dt_From,
            cb.attand_date_to  AS Dt_To,
            cb.present_days    AS Presents,
            cb.file_name       AS File_Name,
            cb.unit_name       As UnitName,
            cb.location_name   AS Location,
            cb.tis_user_name   AS User_Name,
            cb.on_leaves_days  As Leaves,
            cb.lwp_days        As lwp,
            cb.off_days        AS Off_Days,
            cb.gzt_days        As Gzt_Days,
            cb.total_work_hrs  As Work_Hrs,
            cb.ot_hrs          AS Ot_Hrs,
            cb.offdutyhrs      AS OffDutyHrs,
            cb.gztdutyhrs      As GztDutyHrs,
            cb.shorthrs        As Shorts,
            cb.arrears_hrs     AS ArrearsHrs
        FROM
            gtm_consolidated_att_buf cb 
        where 1=1
        AND Transferd='N' 
        AND TRUNC(ATTAND_DATE_FRM)>='{0}'
        AND TRUNC(ATTAND_DATE_TO)<='{1}'
        AND Is_FNF ={2}", DateFrom, DateTo, IsFnf)
        Return Query
    End Function

    Public Function SetPayrollQuery(ByVal DateFrom As String, ByVal DateTo As String, IsFnf As Integer, ByVal EBS As String, ByVal FileName As String) As String
        Dim Query As String = String.Format("Update GTM_CONSOLIDATED_ATT_BUF Set File_Name='{4}' 
        where 1=1
        AND EMPID in  ('{3}')        
        AND Transferd='N' 
        AND TRUNC(ATTAND_DATE_FRM)>='{0}'
        AND TRUNC(ATTAND_DATE_TO)<='{1}'
        AND Is_FNF ={2}", DateFrom, DateTo, IsFnf, EBS, FileName)
        Return Query
    End Function
End Class
