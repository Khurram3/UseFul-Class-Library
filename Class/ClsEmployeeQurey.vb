Public Class ClsEmployeeQurey

    Public Function GetEmployeeData(ByVal EBSCode As String) As String
        Dim Query As String = String.Format("SELECT 
                E.Employee_Code,E.Employee_Name,E.Employee_AppointmentDate AS AppointmentDate, E.Employee_LeftDate AS LeftDate,E.Employee_InOutDate InOutDate ,ESbm.SalaryBreakupMaster_Salary AS Salary,    
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
                WHERE 1=1 
                AND E.Employee_Code IN ('{0}'")
        Return Query
    End Function

    Public Function GetEmployeeForCNICExpire() As String
        Dim Query As String = String.Format("Select Employee_Code, Employee_Name , Employee_NIC ,CNIC_Expire_Date , Employee_LeftDate From Employee  Where  Employee_LeftDate='' or  CAST(Employee_LeftDate as date)  >= DATEADD(DAY, -90, CAST(GETDATE() AS DATE))")
        Return Query
    End Function

    Public Function GetEmployeeCNICFromOra(ByVal EbsCodes As String) As String
        ' Split, trim, wrap in single quotes, and join back
        Dim formattedCodes As String = String.Join(",", EbsCodes.Split(","c).
                                                   Where(Function(code) Not String.IsNullOrWhiteSpace(code)).
                                                   Select(Function(code) "'" & code.Trim() & "'"))
        Dim Query As String = String.Format("SELECT EMPLOYEE_NUMBER, cnic, cnic_expiry FROM gtm_emp_portal_v WHERE EMPLOYEE_NUMBER IN ({0})", formattedCodes)
        Return Query
    End Function

    Public Function InsertQueryOra(ByVal EBSCode) As String
        Dim Query As String = String.Format("Insert Into GetCNICExipreTable (EBSCode) Values ('{0}')", EBSCode)
        Return Query
    End Function

    Public Function GetEmployeeCNICFromOra() As String
        Dim Query As String = String.Format("SELECT EMPLOYEE_NUMBER, cnic, cnic_expiry FROM gtm_emp_portal_v WHERE EMPLOYEE_NUMBER IN (Select EBSCode From GetCNICExipreTable)")
        Return Query
    End Function

    Public Function UpdateCNICExpireDate(ByVal EBSCode As String, CNICExpireDate As String)
        Dim Query As String = String.Format("Update Employee Set CNIC_Expire_Date='{0}' Where Employee_Code='{1}'", CNICExpireDate, EBSCode)
        Return Query
    End Function
End Class
