Imports MySql.Data.MySqlClient


Public Class frmMain
    Dim notifystatus As String = "off"

    Private Sub frmMain_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
       
    End Sub
    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True
        Me.WindowState = FormWindowState.Minimized
        Me.Visible = False
    End Sub

    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Connection.connect_hris()
        Connection.connect_payroll()
        Connection.connect_payroll2()
        RunAtStartup(Application.ProductName, Application.ExecutablePath)

        Me.WindowState = FormWindowState.Minimized
        Me.Visible = False

    End Sub

    Sub sync_cutoff()
        strsql = "SELECT companies.name, cutoff.from_date, cutoff.to_date, occurence, status FROM cutoff, companies WHERE companies.id = cutoff.company_id"
        Dim dt = New DataTable
        da = New MySqlDataAdapter(strsql, con2)
        cmd = New MySqlCommand(strsql, con2)
        da.SelectCommand = cmd
        da.Fill(dt)
        Connection.close_hris()
        strsql = "TRUNCATE tbl_cutoff"
        Connection.exe_payroll()
        For i = 0 To dt.Rows.Count - 1
            Try
                strsql = "INSERT INTO tbl_cutoff(cutoff_range,company_id,occurence_id,from_date,to_date,status) " _
                            & "VALUES('" & CDate(dt.Rows(i)(1).ToString).ToString("d MMM yyyy") & " to " & CDate(dt.Rows(i)(2).ToString).ToString("d MMM yyyy") & "', " _
                            & "'" & dt.Rows(i)(0).ToString & "',(SELECT occurence_id FROM tblref_occurences WHERE name='" & dt.Rows(i)(3).ToString & "'),'" & CDate(dt.Rows(i)(1).ToString).ToString("yyyy-MM-dd") & "','" & CDate(dt.Rows(i)(2).ToString).ToString("yyyy-MM-dd") & "','" & dt.Rows(i)(4).ToString & "')"
                Connection.exe_payroll()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Next
    End Sub

    Sub sync_company()
        strsql = "SELECT name, code FROM companies"
        Dim dt = New DataTable
        da = New MySqlDataAdapter(strsql, con2)
        cmd = New MySqlCommand(strsql, con2)
        da.SelectCommand = cmd
        da.Fill(dt)
        Connection.close_hris()
        strsql = "TRUNCATE tbl_company"
        Connection.exe_payroll()
        For i = 0 To dt.Rows.Count - 1
            Try
                strsql = "INSERT INTO tbl_company(name,code) " _
                            & "VALUES('" & dt.Rows(i)(0).ToString & "','" & dt.Rows(i)(1).ToString & "')"
                Connection.exe_payroll()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Next
    End Sub

    Sub sync_employees()
        strsql = "SELECT emp.id, emp.employee_id, emp.biometric_id, emp.fName, emp.mi, emp.lName, shift.shiftName, " _
                        & "emp.sssNo, emp.phicNo, emp.hdmfNo, emp.taxNo, (com.name) AS company, (bra.name) AS branch, " _
                        & "(pos.name) AS position, rank.rank, taxstat.taxcode, emp.emp_status, serv.basicSalary, emp.lastUpdated " _
                        & "FROM employees emp " _
                        & "LEFT JOIN shiftsgroup shift ON shift.id= emp.shiftgroup_id " _
                        & "LEFT JOIN companies com ON com.id= emp.company_id " _
                        & "LEFT JOIN branches bra ON bra.id= emp.branch_id " _
                        & "LEFT JOIN positions pos ON pos.id= emp.position_id " _
                        & "LEFT JOIN taxstatus taxstat ON taxstat.id= emp.taxstatus_id " _
                        & "LEFT JOIN rank ON rank.id= pos.rank_id " _
                        & "LEFT JOIN services serv ON serv.employee_id= emp.id AND serv.ifcurrent= '1' " _
                        & "WHERE ifNull(emp.employee_id,'') != 'SP-Admin'"
        Dim dt = New DataTable
        da = New MySqlDataAdapter(strsql, con2)
        cmd = New MySqlCommand(strsql, con2)
        da.SelectCommand = cmd
        da.Fill(dt)
        Connection.close_hris()
        For i = 0 To dt.Rows.Count - 1
            'Try
            '    strsql = "REPLACE INTO tbl_employee(id_employee,emp_id,emp_bio_id,fName,mName,lName,shiftgroup,sss_id,phic_id,hdmf_id,tin,company,branch,position,rank,tax_status,employment_status,basic_salary,lastUpdated) " _
            '                        & "VALUES(" & dt.Rows(i)(0).ToString & ",'" & dt.Rows(i)(1).ToString & "'," _
            '                        & "'" & dt.Rows(i)(2).ToString & "','" & dt.Rows(i)(3).ToString & "'," _
            '                        & "'" & dt.Rows(i)(4).ToString & "','" & dt.Rows(i)(5).ToString & "'," _
            '                        & "'" & dt.Rows(i)(6).ToString & "','" & dt.Rows(i)(7).ToString & "'," _
            '                        & "'" & dt.Rows(i)(8).ToString & "','" & dt.Rows(i)(9).ToString & "'," _
            '                        & "'" & dt.Rows(i)(10).ToString & "','" & dt.Rows(i)(11).ToString & "'," _
            '                        & "'" & dt.Rows(i)(12).ToString & "','" & dt.Rows(i)(13).ToString & "'," _
            '                        & "'" & dt.Rows(i)(14).ToString & "','" & dt.Rows(i)(15).ToString & "'," _
            '                        & "'" & dt.Rows(i)(16).ToString & "'," & If(String.IsNullOrEmpty(dt.Rows(i)(17).ToString), 0, dt.Rows(i)(17).ToString) & ",'" & CDate(dt.Rows(i)(18).ToString).ToString("yyyy-MM-dd HH:mm:ss") & "')"
            '    Console.Write(dt.Rows(i)(10).ToString)
            '    Connection.exe_payroll()
            'Catch ex As Exception
            '    MessageBox.Show(ex.ToString)
            'End Try
            'strsql = "SELECT * FROM tbl_employee WHERE id_employee = '" & dt.Rows(i)(0).ToString & "'"
            'cmd = New MySqlCommand(strsql, con)
            'Dim reader1 As MySqlDataReader = cmd.ExecuteReader
            'If reader1.HasRows Then
            '    'read, compare lastupdated and update or not
            '    While reader1.Read()
            '        If CDate(dt.Rows(i)(18).ToString).ToString("yyyy-MM-dd HH:mm:ss") > CDate(reader1(19).ToString).ToString("yyyy-MM-dd HH:mm:ss") Then
            'update
            'Try
            strsql = "UPDATE tbl_employee SET emp_id = '" & dt.Rows(i)(1).ToString & "'," _
                        & "emp_bio_id = '" & dt.Rows(i)(2).ToString & "', fName = '" & dt.Rows(i)(3).ToString & "'," _
                        & "mName = '" & dt.Rows(i)(4).ToString & "', lName = '" & dt.Rows(i)(5).ToString & "'," _
                        & "shiftgroup = '" & dt.Rows(i)(6).ToString & "', sss_id = '" & dt.Rows(i)(7).ToString & "'," _
                        & "phic_id = '" & dt.Rows(i)(8).ToString & "', hdmf_id = '" & dt.Rows(i)(9).ToString & "'," _
                        & "tin = '" & dt.Rows(i)(10).ToString & "', company = '" & dt.Rows(i)(11).ToString & "'," _
                        & "branch = '" & dt.Rows(i)(12).ToString & "', position = '" & dt.Rows(i)(13).ToString & "'," _
                        & "rank = '" & dt.Rows(i)(14).ToString & "', tax_status = '" & dt.Rows(i)(15).ToString & "'," _
                        & "employment_status = '" & dt.Rows(i)(16).ToString & "', basic_salary = " & If(String.IsNullOrEmpty(dt.Rows(i)(17).ToString), 0, dt.Rows(i)(17).ToString) & "," _
                        & "lastUpdated = '" & CDate(dt.Rows(i)(18).ToString).ToString("yyyy-MM-dd HH:mm:ss") & "' WHERE id_employee =" & dt.Rows(i)(0).ToString
            'Console.Write(StrSql)
            'QryReadP()
            'cmd.ExecuteNonQuery()
            Connection.exe_payroll()

            strsql = "SELECT * FROM tbl_employee WHERE id_employee = '" & dt.Rows(i)(0).ToString & "'"
            Connection.query_payroll()
            If Not dr.HasRows Then
                'While dr.Read
                strsql = "INSERT INTO tbl_employee(id_employee,emp_id,emp_bio_id,fName,mName,lName,shiftgroup,sss_id,phic_id,hdmf_id,tin,company,branch,position,rank,tax_status,employment_status,basic_salary,lastUpdated) " _
                    & "VALUES(" & dt.Rows(i)(0).ToString & ",'" & dt.Rows(i)(1).ToString & "'," _
                    & "'" & dt.Rows(i)(2).ToString & "','" & dt.Rows(i)(3).ToString & "'," _
                    & "'" & dt.Rows(i)(4).ToString & "','" & dt.Rows(i)(5).ToString & "'," _
                    & "'" & dt.Rows(i)(6).ToString & "','" & dt.Rows(i)(7).ToString & "'," _
                    & "'" & dt.Rows(i)(8).ToString & "','" & dt.Rows(i)(9).ToString & "'," _
                    & "'" & dt.Rows(i)(10).ToString & "','" & dt.Rows(i)(11).ToString & "'," _
                    & "'" & dt.Rows(i)(12).ToString & "','" & dt.Rows(i)(13).ToString & "'," _
                    & "'" & dt.Rows(i)(14).ToString & "','" & dt.Rows(i)(15).ToString & "'," _
                    & "'" & dt.Rows(i)(16).ToString & "'," & If(String.IsNullOrEmpty(dt.Rows(i)(17).ToString), 0, dt.Rows(i)(17).ToString) & ",'" & CDate(dt.Rows(i)(18).ToString).ToString("yyyy-MM-dd HH:mm:ss") & "')"
                Connection.exe_payroll2()
                'End While
            End If
            Connection.close_payroll()
            'Catch e As MySqlException
            'MessageBox.Show(e.ToString)
            'End Try
            '    End If
            'End While
            'Else
            '    'insert
            '    Try
            '        strsql = "INSERT INTO tbl_employee(id_employee,emp_id,emp_bio_id,fName,mName,lName,shiftgroup,sss_id,phic_id,hdmf_id,tin,company,branch,position,rank,tax_status,employment_status,basic_salary,lastUpdated) " _
            '                    & "VALUES(" & dt.Rows(i)(0).ToString & ",'" & dt.Rows(i)(1).ToString & "'," _
            '                    & "'" & dt.Rows(i)(2).ToString & "','" & dt.Rows(i)(3).ToString & "'," _
            '                    & "'" & dt.Rows(i)(4).ToString & "','" & dt.Rows(i)(5).ToString & "'," _
            '                    & "'" & dt.Rows(i)(6).ToString & "','" & dt.Rows(i)(7).ToString & "'," _
            '                    & "'" & dt.Rows(i)(8).ToString & "','" & dt.Rows(i)(9).ToString & "'," _
            '                    & "'" & dt.Rows(i)(10).ToString & "','" & dt.Rows(i)(11).ToString & "'," _
            '                    & "'" & dt.Rows(i)(12).ToString & "','" & dt.Rows(i)(13).ToString & "'," _
            '                    & "'" & dt.Rows(i)(14).ToString & "','" & dt.Rows(i)(15).ToString & "'," _
            '                    & "'" & dt.Rows(i)(16).ToString & "'," & If(String.IsNullOrEmpty(dt.Rows(i)(17).ToString), 0, dt.Rows(i)(17).ToString) & ",'" & CDate(dt.Rows(i)(18).ToString).ToString("yyyy-MM-dd HH:mm:ss") & "')"
            '        'Console.Write(StrSql)
            '        'QryReadP()
            '        'cmd.ExecuteNonQuery()
            '        Connection.exe_payroll()
            '    Catch e As MySqlException
            '        MessageBox.Show(e.ToString)
            '    End Try
            'End If
            'reader1.Dispose()
        Next
    End Sub
    Sub sync_leaves()
         StrSql = "SELECT leaveapp.id, leaveapp.employee_id, " _
                       & "leaves.name AS 'Leave Type', leaveapp.durFrom AS 'From Date', leaveapp.durTo AS 'To Date', leaveapp.dateFiled AS 'Date Filed', " _
                       & "leaveapp.days_applied AS 'Days Applied', leaveapp.mode, leaveapp.reason AS 'Reason', leaveapp.status AS 'Status' FROM leaveapp, leaves, employees " _
                       & "WHERE leaveapp.leave_id = leaves.id AND leaveapp.employee_id = employees.id AND leaveapp.status = 'Approved by HR'"
        Dim dt = New DataTable
        da = New MySqlDataAdapter(strsql, con2)
        cmd = New MySqlCommand(strsql, con2)
        da.SelectCommand = cmd
        da.Fill(dt)
        Connection.close_hris()
        For i = 0 To dt.Rows.Count - 1
            Try
                StrSql = "REPLACE INTO tbl_leaves(id, employee_id,leave_type,durFrom,durTo,dateFiled,mode,days_applied,reason,status) " _
                                & "VALUES(" & dt.Rows(i)(0).ToString & "," & dt.Rows(i)(1).ToString & ",'" _
                                & dt.Rows(i)(2).ToString & "','" & CDate(dt.Rows(i)(3).ToString).ToString("yyyy-MM-dd") & "','" _
                                & CDate(dt.Rows(i)(4).ToString).ToString("yyyy-MM-dd") & "','" & CDate(dt.Rows(i)(5).ToString).ToString("yyyy-MM-dd") & "','" _
                                & dt.Rows(i)(6).ToString & "','" & dt.Rows(i)(7).ToString & "','" _
                                & dt.Rows(i)(8).ToString & "','" & dt.Rows(i)(9).ToString & "')"
                Connection.exe_payroll()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Next
    End Sub

    Sub sync_loans()
        strsql = "SELECT loans.id, employees.id as employee_id, loantype.loantype AS 'Loan', " _
                        & "lendingcompany.name AS 'Lending Company', loans.amount AS 'Amount', loans.term AS 'Term', " _
                        & "loans.monthlyAmortization AS 'Monthly Amortization', loans.startDate AS 'From', " _
                        & "loans.endDate AS 'To', loans.remarks AS 'Remarks' " _
                        & "FROM loans, loantype, employees, lendingcompany WHERE loans.employee_id = employees.id " _
                        & "AND lendingcompany.id = loans.lendingCompany_id AND loantype.id = loans.loantype_id"
        Dim dt = New DataTable
        da = New MySqlDataAdapter(strsql, con2)
        cmd = New MySqlCommand(strsql, con2)
        da.SelectCommand = cmd
        da.Fill(dt)
        Connection.close_hris()
        For i = 0 To dt.Rows.Count - 1
            Try
                strsql = "REPLACE INTO tbl_loans(loan_id,employee_id,loan_type,lendingCompany,amount,term,monthlyAmortization,startDate,endDate,remarks)" _
                                 & "VALUES(" & dt.Rows(i)(0).ToString & "," & dt.Rows(i)(1).ToString & ",'" & dt.Rows(i)(2).ToString & "','" _
                                 & dt.Rows(i)(3).ToString & "'," & dt.Rows(i)(4).ToString & ",'" _
                                 & dt.Rows(i)(5).ToString & "'," & dt.Rows(i)(6).ToString & ",'" _
                                 & CDate(dt.Rows(i)(7).ToString).ToString("yyyy-MM-dd") & "','" & CDate(dt.Rows(i)(8).ToString).ToString("yyyy-MM-dd") & "','" _
                                 & dt.Rows(i)(9).ToString & "')"
                Connection.exe_payroll()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Next
    End Sub
    Sub sync_overtime()
         StrSql = "SELECT overtime.id, overtime.employee_id, " _
                       & "overtime.reason AS 'Reason', overtime.dateFiled AS 'Date Filed', " _
                       & "overtime.dateRequested AS 'Date Requested', overtime.timeStart AS 'From', " _
                       & "overtime.timeEnd AS 'To', overtime.totalHours AS 'Total Hours', overtime.status AS 'Status' " _
                       & "FROM overtime WHERE overtime.status = 'Approved by HR'"
        Dim dt = New DataTable
        da = New MySqlDataAdapter(strsql, con2)
        cmd = New MySqlCommand(strsql, con2)
        da.SelectCommand = cmd
        da.Fill(dt)
        Connection.close_hris()
        For i = 0 To dt.Rows.Count - 1
            Try
                strsql = "REPLACE INTO tbl_overtime(id,employee_id,reason,dateFiled,dateRequested,timeStart,timeEnd,totalHours,status)" _
                                & "VALUES(" & dt.Rows(i)(0).ToString & "," & dt.Rows(i)(1).ToString & ",'" _
                                & dt.Rows(i)(2).ToString & "','" & CDate(dt.Rows(i)(3).ToString).ToString("yyyy-MM-dd") & "','" _
                                & CDate(dt.Rows(i)(4).ToString).ToString("yyyy-MM-dd") & "','" & dt.Rows(i)(5).ToString & "','" _
                                & dt.Rows(i)(6).ToString & "','" & dt.Rows(i)(7).ToString & "','" _
                                & dt.Rows(i)(8).ToString & "')"
                Connection.exe_payroll()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Next
    End Sub
    Sub sync_shifts()
       StrSql = "SELECT shifts.id, shifts.day as 'Day', shifts.timein as 'From', " _
                         & "shifts.timeout as 'To', shiftsgroup.shiftName as 'Shift Name' " _
                         & "FROM shifts, shiftsgroup WHERE shifts.shiftgroup_id = shiftsgroup.id"
        Dim dt = New DataTable
        da = New MySqlDataAdapter(strsql, con2)
        cmd = New MySqlCommand(strsql, con2)
        da.SelectCommand = cmd
        da.Fill(dt)
        Connection.close_hris()
        For i = 0 To dt.Rows.Count - 1
            Try
                StrSql = "REPLACE INTO tbl_shifts(id,day,timein,timeout,shiftgroup)" _
                                & "VALUES('" & dt.Rows(i)(0).ToString & "','" & dt.Rows(i)(1).ToString & "','" & dt.Rows(i)(2).ToString & "','" & dt.Rows(i)(3).ToString & "','" & dt.Rows(i)(4).ToString & "')"
                Connection.exe_payroll()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Next
    End Sub
    Sub sync_allowance()
        strsql = "SELECT sa.id, sa.employee_id, (al.name) AS allowance, sa.amount FROM serviceallowance sa " _
                        & "JOIN allowances al ON al.id= sa.allowance_id LEFT JOIN services svs ON svs.id = sa.service_id " _
                        & "WHERE svs.ifcurrent = '1'"
        Dim dt = New DataTable
        da = New MySqlDataAdapter(strsql, con2)
        cmd = New MySqlCommand(strsql, con2)
        da.SelectCommand = cmd
        da.Fill(dt)
        Connection.close_hris()
        strsql = "TRUNCATE tbl_allowances"
        Connection.exe_payroll()
        For i = 0 To dt.Rows.Count - 1
            Try
                strsql = "INSERT INTO tbl_allowances(employee_id,name,amount)" _
                               & "VALUES('" & dt.Rows(i)(1).ToString & "','" & dt.Rows(i)(2).ToString & "','" & dt.Rows(i)(3).ToString & "')"
                Connection.exe_payroll()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Next
    End Sub
    Private Sub frmMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            NotifyIcon1.Visible = True
            'NotifyIcon1.Icon = SystemIcons.Application
            NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
            NotifyIcon1.BalloonTipTitle = "HRIS"
            NotifyIcon1.BalloonTipText = "Synching files...."
            NotifyIcon1.ShowBalloonTip(50000)
            'Me.Hide()
            ShowInTaskbar = False
        End If
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        sync_company()
        sync_cutoff()
        sync_employees()
        sync_allowance()
        sync_shifts()
        sync_overtime()
        sync_loans()
        sync_leaves()
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        NotifyIcon1.BalloonTipText = e.ProgressPercentage.ToString() + "%"
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If notifystatus = "off" Then
            notifystatus = "on"
            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        notifystatus = "off"
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        ShowInTaskbar = True
        Me.WindowState = FormWindowState.Normal
        NotifyIcon1.Visible = False
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
End Class
