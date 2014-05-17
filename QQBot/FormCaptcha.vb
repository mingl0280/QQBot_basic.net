Imports System.Net
Imports System.IO

Public Class FormCaptcha

    Private Sub FormCaptcha_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button1.Enabled = False
        FlushCaptcha()
    End Sub

    Private Function FlushCaptcha()
        Dim captchalink As String = "http://captcha.qq.com/getimage?&uin=" + LoginInfo.SourceID + "&aid=10020101&0." + GetRandomNum(16)
        Dim webdownloader As WebClient = New WebClient
        Dim bbyte() As Byte
        bbyte = webdownloader.DownloadData(captchalink)

        Dim CookieKeys As String() = webdownloader.ResponseHeaders.AllKeys
        Dim CookieGet As String = ""
        For i As Integer = 0 To UBound(CookieKeys)
            If CookieKeys(i).Equals(sCookie) Then
                CookieGet = webdownloader.ResponseHeaders.Get(i)
            End If
        Next
        LoginInfo.CapCookie = CookieGet
        Dim ms As MemoryStream = New MemoryStream()
        ms.Write(bbyte, 0, bbyte.Length)
        ms.Flush()
        PictureBox1.Image = Image.FromStream(ms)
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then Exit Sub
        LoginInfo.key = TextBox1.Text
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            Button1.Enabled = False
        Else
            Button1.Enabled = True
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        FlushCaptcha()
        TextBox1.Text = ""
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
End Class