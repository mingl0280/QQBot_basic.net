Imports System.Threading

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MessageBox.Show("请输入用户名与密码", "警告", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
        Label4.Visible = True
        Dim hashpass As String = MD5_Encrypt(TextBox2.Text)
        LoginInfo.QPass = hashpass
        LoginInfo.SourceID = TextBox1.Text
        getLoginSig()
        'sig validation
        getLoginCookiesAndKey()
        'cookie and key validation
        LoginInfo.QPass = getNewPassHash_o()
        Dim SendCookie As String = ""
        If LoginInfo.CapCookie = "" Then
            SendCookie = ClearifyCookieStr(LoginInfo.VerifyCookie)
        Else
            SendCookie = ClearifyCookieStr(LoginInfo.VerifyCookie + LoginInfo.CapCookie)
        End If
        LoginSSLPage(SendCookie)

        'login successful check

        'if not successful then 

        'login successful check over
        Dim webURL As String = LoginInfo.LoginCompleteParametersArray(2)



        Dim uin As String = LoginInfo.LoginCompleteCookie.Item("uin").Value
        Dim ptwebqq As String = LoginInfo.LoginCompleteCookie.Item("ptwebqq").Value
        Dim skey As String = LoginInfo.LoginCompleteCookie.Item("skey").Value

        uin = uin.Substring(1, uin.Length - 1)

        Label4.Visible = False

    End Sub

    Private Sub initLogFile()
        Dim fname As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        fname += "\sysmon"
        If IO.File.Exists(fname) = False Then
            My.Computer.FileSystem.CreateDirectory(fname)
            Thread.Sleep(200)
        End If
        Dim ttime As Date = Now
        With ttime
            fname += "\Logging_" + .Year.ToString + "-" _
                + .Month.ToString + "-" _
                + .Day.ToString + "_" _
                + .Hour.ToString + "." _
                + .Minute.ToString + "." _
                + .Second.ToString + ".log"
        End With
        LogFileStream = IO.File.Create(fname)
        Thread.Sleep(200)
        LogFile = fname
        WriteToLog(LOGTYPE.TYPE_INFO, "Logging file created.")
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label4.Visible = False
    End Sub
End Class
