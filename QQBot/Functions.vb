Imports System.Security
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Encoding
Imports System.Threading
Imports System.Web
Imports System.Net.WebSockets
Imports System.Net
Imports System.IO

#Const REMOVE = "REM"


Module Functions

    Public Function getMD5String(ByVal s As String) As String
        Dim bbyte() As Byte = ASCII.GetBytes(s)
        Dim bbyte_hashed() As Byte
        Dim md5trans As MD5CryptoServiceProvider = New MD5CryptoServiceProvider()
        md5trans.ComputeHash(bbyte)
        bbyte_hashed = md5trans.Hash
        Dim ctstr As String = ""
        ctstr = BitConverter.ToString(bbyte_hashed)
        ctstr = ctstr.Replace("-", "")
        Return ctstr.ToUpper()
    End Function

    Public Function getLoginSig()
        Dim webpage As String = "https://ui.ptlogin2.qq.com/cgi-bin/login?daid=164&target=self&style=5&mibao_css=m_webqq&appid=1003903&enable_qlogin=0&no_verifyimg=1&s_url=http%3A%2F%2Fweb2.qq.com%2Floginproxy.html&f_url=loginerroralert&strong_login=1&login_state=10&t=20130903001"

        Dim webDownloader As WebClient = New WebClient()
        Dim bbyte() As Byte
        Dim webdocument As String
        Try
            bbyte = webDownloader.DownloadData(webpage)
            webdocument = UTF8.GetString(bbyte)
        Catch ex As Exception
            Exit Function
        End Try
        Dim i, j As Integer
        i = webdocument.IndexOf("var g_login_sig=encodeURIComponent(""") + "var g_login_sig=encodeURIComponent(""".Length
        j = webdocument.IndexOf(""");", i)
        LoginInfo.sig = webdocument.Substring(i, j - i)
    End Function

    Public Function getLoginCookiesAndKey()
        Dim webDownloader As WebClient = New WebClient
        webDownloader.Headers.Add("Content-Type", "application/x-www-form-urlencoded")
        webDownloader.Headers.Add("UserAgent", "Mozilla/5.0(iPad; U; CPU iPhone OS 3_2 like Mac OS X; en-us) AppleWebKit/531.21.10 (KHTML, like Gecko) Version/4.0.4 Mobile/7B314 Safari/531.21.10)")
        Dim RequestURL As String = "https://ssl.ptlogin2.qq.com/check?uin="
        Dim RequestData As String = LoginInfo.SourceID + _
                                    "&appid=1003903&js_ver=10052&js_type=0&login_sig=" + _
                                    LoginInfo.sig + _
                                    "&u1=http%3A%2F%2Fweb2.qq.com%2Floginproxy.html&r=0." + _
                                    GetRandomNum(16)
        Dim innerHTML As String = webDownloader.UploadString(RequestURL + RequestData, "POST", "")
        Dim CookieKeys As String() = webDownloader.ResponseHeaders.AllKeys
        Dim CookieGet As String = ""
        For i As Integer = 0 To UBound(CookieKeys)
            If CookieKeys(i).Equals(sCookie) Then
                CookieGet = webDownloader.ResponseHeaders.Get(i)
            End If
        Next
        LoginInfo.VerifyCookie = CookieGet
        Dim paramList As String() = innerHTML.Split(",")
        If paramList.Length <> 3 Then
            Exit Function
        Else
            LoginInfo.newRandStr = paramList(2).Replace("'", "").Replace(");", "")
            LoginInfo.VerifyCookie = CookieGet
            LoginInfo.key = paramList(1).Replace("'", "")
            paramList(0) = paramList(0).Replace("ptui_checkVC(", "")
            If Not ((paramList(0) = "'0'" And LoginInfo.key.Length = 4) Or (paramList(0) = "'1'" And LoginInfo.key.Length > 4)) Then
                Exit Function
            ElseIf (paramList(0) = "'1'" And LoginInfo.key.Length > 4) Then
                If onValidationCodeShown() = "INVALID_CAP" Then
                    Exit Function
                End If
            End If
        End If
    End Function

    Function onValidationCodeShown() As String
        Dim ret As DialogResult = FormCaptcha.ShowDialog()
        If ret = DialogResult.Cancel Then
            Return "INVALID_CAP"
        End If
    End Function

    Public Function ClearifyCookieStr(ByRef param1 As String) As String
        Dim dststr As String = ""
        Dim tempstrarr As String()
        tempstrarr = param1.Split(";")
        For i As Integer = 0 To UBound(tempstrarr)
            If tempstrarr(i).IndexOf("PATH") >= 0 Or tempstrarr(i).IndexOf("DOMAIN") >= 0 Then Continue For
            dststr += tempstrarr(i) + ";"
        Next
        Dim counter As Integer = 0
        For j As Integer = UBound(tempstrarr) To 0
            If counter = 2 Then Exit For
        Next
        dststr.Replace(",", "")
        dststr.Remove(dststr.Length - 1, 1)
        param1 = dststr
        Return dststr
    End Function

    Public Function getNewPassHash_o() As String
        Dim combstr As String = LoginInfo.QPass + LoginInfo.newRandStr.Replace("\x", "")
        Dim bbyte As Byte() = hexstr2byte(combstr)
        Dim newHash1 As String = MD5_Encrypt(bbyte)
        Dim finalresult As String = MD5_Encrypt(newHash1.ToUpper + LoginInfo.key.ToUpper)
        Return finalresult
    End Function

    Private Function hexstr2byte(ByVal s As String) As Byte()
        Dim xbyte(Math.Floor(s.Length / 2) - 1) As Byte
        For i As Integer = 0 To s.Length - 1 Step 2
            xbyte(i / 2) = CInt("&h" + s.Substring(i, 2))
        Next
        Return xbyte
    End Function

    Public Function GetRandomNum(ByVal numlength As Integer) As String
        Dim k As String = ""
        Dim r As Random = New Random()
        For i As Integer = 0 To numlength - 1
            k += r.Next(9).ToString
        Next
        Return k
    End Function

    Public Function LoginSSLPage(ByVal cookiestr As String)
        Dim sigstr As String = ""
        If LoginInfo.sig <> "" Then sigstr = "&login_sig=" + LoginInfo.sig
        Dim url = "https://ssl.ptlogin2.qq.com/login?u=" + LoginInfo.SourceID + "&p=" + LoginInfo.QPass + _
            "&verifycode=" + LoginInfo.key + _
            "&webqq_type=10&remember_uin=1&login2qq=1&aid=1003903&u1=http%3A%2F%2Fweb2.qq.com%2Floginproxy.html%3Flogin2qq%3D1%26webqq_type%3D10&h=1&ptredirect=0&ptlang=2052&daid=164&from_ui=1&pttype=1&dumy=&fp=loginerroralert&action=" + _
            GetActionCode() + _
            "&mibao_css=m_webqq&t=1&g=1&js_type=0&js_ver=10052" + sigstr
        Dim webReq As HttpWebRequest
        webReq = HttpWebRequest.Create(url)
        With webReq
            SetCookieHeaders(webReq, cookiestr)
            .ContentType = "application/x-www-form-urlencoded"
            .Headers.Add("Accept-Language", "zh-cn")
            .Method = "GET"
        End With
        Dim StreamData As Stream
        Dim webResp As HttpWebResponse = webReq.GetResponse()
        StreamData = webResp.GetResponseStream
        Dim nCook = webResp.Cookies
        Dim innerHTML = New StreamReader(StreamData, Encoding.GetEncoding("UTF-8")).ReadToEnd()
        StreamData.Close()
        Dim startP As Integer = innerHTML.IndexOf("ptuiCB(") + "ptuiCB(".Length
        Dim endP As Integer = innerHTML.IndexOf(")")
        Dim paramstr As String = innerHTML.Substring(startP, endP - startP)
        Dim paramArr As String()
        paramArr = paramstr.Split(",")
        For i As Integer = 0 To UBound(paramArr)
            paramArr(i).Trim("'")
        Next
        LoginInfo.LoginCompleteCookie = nCook
        LoginInfo.LoginCompleteParametersArray = paramArr

    End Function

    Public Function GetActionCode() As String
        Dim r As Random = New Random
        Dim retstr As String
        retstr = r.Next(1, 8).ToString + "-" + r.Next(1, 20).ToString + "-" + r.Next(10000, 30000).ToString
        Return retstr
    End Function

    Public Function SetCookieHeaders(ByRef hClient As HttpWebRequest, ByVal CookieString As String)
        Dim CookieParams() As String = CookieString.Split(";")
        hClient.CookieContainer = New CookieContainer()
        Dim addr As New Uri("http://ptlogin2.qq.com/")
        Dim addr2 As New Uri("http://www.qq.com/")
        Dim addr3 As New Uri("https://ssl.ptlogin2.qq.com/")
        For i As Integer = 0 To UBound(CookieParams)
            Dim CookItem() As String = CookieParams(i).Split("=")
            Dim CK As Cookie
            If UBound(CookItem) >= 1 Then
                CK = New Cookie(CookItem(0).Replace(",", ""), CookItem(1).Replace(",", ""))
                hClient.CookieContainer.Add(addr, CK)
                hClient.CookieContainer.Add(addr2, CK)
                hClient.CookieContainer.Add(addr3, CK)
            End If
        Next
    End Function

    Public Function readini(ByVal key As String, ByRef dststr As String, Optional ByVal ReadLocal As Boolean = True) As String
        Dim isex As Integer
        isex = IO.File.Exists("Config.ini")
        If Not isex Then
            'IO.File.Create(Application.StartupPath + "\Config.ini")
            Return ""
        Else
            Dim str As String
            str = ""
            str = LSet(str, 512)
            Dim currentdir As String = ""
            If ReadLocal = False Then
                currentdir = NetINIFile
            Else
                currentdir = Environment.CurrentDirectory + "\Config.ini"
            End If
            GetPrivateProfileString("main", key, "", str, Len(str), currentdir)
            dststr = Left(str, InStr(str, Chr(0)) - 1)
            Return 1
        End If
    End Function

    '写入ini函数
    Public Function writeini(ByVal key As String, ByVal str As String) As Boolean
        Dim isex As Integer
        isex = File.Exists("Config.ini")
        If Not isex Then
            File.Create(Environment.CurrentDirectory + "\Config.ini")
            Thread.Sleep(500)
        End If
        Dim path As String
        path = Application.StartupPath + "\Config.ini"
        WritePrivateProfileString("main", key, str, path)
    End Function

    'Base64加密与解密函数
    Public Function Base64Encode(ByVal s_src As String) As String
        Dim bbyte() As Byte
        bbyte = ASCII.GetBytes(s_src)
        Dim s_enc As String = System.Convert.ToBase64String(bbyte)
        Return s_enc
    End Function

    Public Function Base64Decode(ByVal s_enc As String) As String
        Dim bbyte() As Byte
        bbyte = System.Convert.FromBase64String(s_enc)
        Dim s_src As String = ASCII.GetString(bbyte)
        Return s_src
    End Function

    Public Function WriteToLog(ByVal LogType As LOGTYPE, ByVal LogInfo As String)
#If DEBUG Then
        Debug.WriteLine(LogInfo)
#End If
        Dim ttime As Date = New Date
        ttime = Now
        Dim timestr As String = "[ " + ttime.ToShortDateString + " " + ttime.ToShortTimeString + ":" + ttime.Second.ToString + " ] "
        Dim eventTypeStr As String = ""
        Select Case LogType
            Case LogType.TYPE_ERROR
                eventTypeStr = "- ERROR - "
            Case LogType.TYPE_INFO
                eventTypeStr = "- INFO - "
            Case LogType.TYPE_WARNING
                eventTypeStr = "- WARNING - "
        End Select
        Dim w As StreamWriter = New StreamWriter(LogFileStream)
        Dim bbyte() As Byte
        Dim UTFStr As String
        bbyte = ASCII.GetBytes(timestr + eventTypeStr)
        UTFStr = UTF8.GetString(bbyte)
        w.WriteLine(UTFStr + LogInfo)
        w.Flush()
    End Function

End Module
