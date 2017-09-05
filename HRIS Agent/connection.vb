
Imports MySql.Data.MySqlClient
Imports System.Data.Odbc
Imports System.Environment
Imports System.Net
Imports System.IO

Module Connection
    Public con, con2, con3 As New MySqlConnection
    Public cmd, cmd2, cmd3 As New MySqlCommand
    Public dr, dr2 As MySqlDataReader
    Public da As New MySqlDataAdapter
    Public ds As New DataSet
    Public dt As New DataTable
    Public strsql As String
    Public ini_file As String
    Public payroll_host, payroll_user, payroll_pw, payroll_db, payroll_port As String
    Public hris_host, hris_user, hris_pw, hris_db, hris_port As String
    Public islogin As Boolean
    Public user_role As Integer

    Public Sub Main()
        Application.EnableVisualStyles()                        ' This is already default on Visual Basic Application 
        Application.SetCompatibleTextRenderingDefault(False)    ' This is already default on Visual Basic Application 

        ini_file = Application.StartupPath & "\agent.ini"           'the main configuration .ini file
        payroll_db = ReadINI("payroll_Database", "payroll_db", ini_file)
        payroll_user = ReadINI("payroll_Database", "payroll_user", ini_file)
        payroll_pw = ReadINI("payroll_Database", "payroll_pw", ini_file)
        payroll_host = ReadINI("payroll_Database", "payroll_host", ini_file)

        hris_db = ReadINI("hris_Database", "hris_db", ini_file)
        hris_user = ReadINI("hris_Database", "hris_user", ini_file)
        hris_pw = ReadINI("hris_Database", "hris_pw", ini_file)
        hris_host = ReadINI("hris_Database", "hris_host", ini_file)
        If (Trim(payroll_db) = "") Or (Trim(hris_db) = "") Then
            MsgBox("Please check config.ini file or please ask administrator.", vbCritical)
        Else
            'Application.Run(New login())  ' Method use to execute run command on what form to display first on the Application
            'If islogin Then
            Application.Run(New frmMain())
            'End If
        End If
    End Sub

    Sub connect_payroll()
        's_host = "localhost"
        's_user = "root"
        's_pw = ""
        's_db = "meprs"
        If con.State = ConnectionState.Open Then
            con.Close()
        End If
        con.ConnectionString = "server=" & payroll_host & ";" _
                             & "user id=" & payroll_user & ";" _
                             & "password=" & payroll_pw & ";" _
                             & "database=" & payroll_db & ";"
        Try
            con.Open()
        Catch ex As Exception
            con.Close()
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Error Connection!")
        End Try
    End Sub
    Sub connect_payroll2()
        's_host = "localhost"
        's_user = "root"
        's_pw = ""
        's_db = "meprs"
        If con3.State = ConnectionState.Open Then
            con3.Close()
        End If
        con3.ConnectionString = "server=" & payroll_host & ";" _
                             & "user id=" & payroll_user & ";" _
                             & "password=" & payroll_pw & ";" _
                             & "database=" & payroll_db & ";"
        Try
            con3.Open()
        Catch ex As Exception
            con3.Close()
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Error Connection!")
        End Try
    End Sub
    Sub connect_hris()
        's_host = "localhost"
        's_user = "root"
        's_pw = ""
        's_db = "meprs"
        If con2.State = ConnectionState.Open Then
            con2.Close()
        End If
        con2.ConnectionString = "server=" & hris_host & ";" _
                             & "user id=" & hris_user & ";" _
                             & "password=" & hris_pw & ";" _
                             & "database=" & hris_db & ";"
        Try
            con2.Open()
        Catch ex As Exception
            con2.Close()
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Error Connection!")
        End Try
    End Sub
    Sub query_payroll()
        cmd = New MySqlCommand(strsql, con)
        cmd.ExecuteNonQuery()
        dr = cmd.ExecuteReader
    End Sub

    Sub close_payroll()
        cmd.Dispose()
        dr.Close()
    End Sub
    
    Sub exe_payroll()
        cmd = New MySqlCommand(strsql, con)
        cmd.ExecuteNonQuery()
        cmd.Dispose()
    End Sub
    Sub exe_payroll2()
        cmd3 = New MySqlCommand(strsql, con3)
        cmd3.ExecuteNonQuery()
        cmd3.Dispose()
    End Sub
    Sub query_hris()
        cmd2 = New MySqlCommand(strsql, con2)

        cmd2.ExecuteNonQuery()
        dr2 = cmd2.ExecuteReader
    End Sub

    Sub close_hris()
        cmd2.Dispose()
        'dr2.Close()
        da.Dispose()
    End Sub
    Sub exe_hris()
        cmd2 = New MySqlCommand(strsql, con2)
        cmd2.ExecuteNonQuery()
        cmd2.Dispose()
    End Sub

    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Int32
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer

    Function ReadINI(ByVal pSection As String, ByVal pKey As String, ByVal pIniFilename As String) As String
        Const DEFAULT_VALUE As String = ""          'default value or an empty string
        Dim lngReturnValue As Long                 'return value of the API call
        Dim strResult As String                     'the resulting string
        Dim lngBuffer As Long
        'length of the resulting string

        'If (Trim$(pSection) = "") Then MsgBox("DEBUG::mod_IniParser::ReadINI() - Bad or missing parameter pSection") : Exit Function
        'If (Trim$(pKey) = "") Then MsgBox("DEBUG::mod_IniParser::ReadINI() - Bad or missing parameter pKey") : Exit Function
        'If (Trim$(pIniFilename) = "") Then MsgBox("DEBUG::mod_IniParser::ReadINI() - Bad or missing parameter pIniFilename") : Exit Function

        strResult = StrDup(1000, vbNullChar)         'pad the resulting string with NULL chars
        lngBuffer = System.Text.Encoding.Unicode.GetByteCount(strResult)                 'get the length of the resulting string

        lngReturnValue = GetPrivateProfileString(pSection, pKey, DEFAULT_VALUE, strResult, lngBuffer, pIniFilename)

        'remove comment
        If InStr(strResult, vbTab & ";", CompareMethod.Text) > 0 Then
            strResult = Trim$(Left$(strResult, InStr(strResult, vbTab & ";", CompareMethod.Text)))
        End If

        strResult = Replace(strResult, vbNullChar, "", , , CompareMethod.Text)       'strip-off all NULL characters

        ReadINI = Trim$(Replace(strResult, vbTab, "", , , CompareMethod.Text))       'strip-off all TAB characters

    End Function

    Function SaveINI(ByVal pSection As String, ByVal pKey As String, ByVal pValue As String, ByVal pIniFilename As String) As Long


        Dim lngReturnValue As Long

        'If (Trim$(pSection) = "") Then MsgBox("DEBUG::mod_INIParser::SaveINI() - Bad or missing parameter pSection") : Exit Function
        'If (Trim$(pKey) = "") Then MsgBox("DEBUG::mod_INIParser::SaveINI() - Bad or missing parameter pKey") : Exit Function
        'If (Trim$(pIniFilename) = "") Then MsgBox("DEBUG::mod_INIParser::ReadINI() - Bad or missing parameter pIniFilename") : Exit Function

        'Comment = ReadComment(pSection, pKey, pValue)

        lngReturnValue = WritePrivateProfileString(pSection, pKey, pValue, pIniFilename)

        SaveINI = lngReturnValue

    End Function
    Sub save_log(ByVal username As String)
        SaveSetting("NDCRegSyS", "Logs", "last_user", username)
    End Sub
End Module
