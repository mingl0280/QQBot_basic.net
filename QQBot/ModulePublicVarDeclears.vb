Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Text.Encoding


Module ModulePublicVarDeclears
#Region "Declare APIs"
    '两个读写ini文件的函数
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, _
         ByVal lpKeyName As String, _
         ByVal lpDefault As String, _
         ByVal lpReturnedString As String, _
         ByVal nSize As Int32, _
         ByVal lpFileName As String) _
     As Int32
    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, _
         ByVal lpKeyName As String, _
         ByVal lpString As String, _
         ByVal lpFileName As String) _
     As Boolean
#End Region

#Region "Declare Vars"

    Public NetINIFile As String = ""
    Public NetAuthExe As String = ""
    Public NetIPAddr As String = ""
    Enum LOGTYPE
        TYPE_INFO
        TYPE_WARNING
        TYPE_ERROR
    End Enum

    Public LoginInfo As LoginInformation
    Public LogFileStream As FileStream
    Public LogFile As String

    Public Const sCookie As String = "Set-Cookie"

    ''' <summary>
    ''' 登录信息存储结构
    ''' <param name="SourceID">QQ号</param>
    ''' <param name="QPass">一次哈希的QQ密码</param>
    ''' <param name="VerifyCookie">登录页Cookie</param>
    ''' <param name="newRandStr">新随机码</param>
    ''' <param name="key">验证码</param>
    ''' <param name="LoginCompleteCookie">登录完毕后的页面Cookie</param>
    ''' <param name="LoginCompleteParametersArray">登录完毕后的页面参数列表</param>
    ''' </summary>
    Structure LoginInformation
        Dim sig As String
        Dim key As String
        Dim SourceID As String 'QQ号
        Dim QPass As String '哈希QQ密码
        Dim VerifyCookie As String '登录Cookie
        Dim ClientID As String
        Dim newRandStr As String '新随机码
        Dim CapStr As String '验证码
        Dim CapCookie As String '验证码页面Cookie
        Dim LoginCompleteCookie As CookieCollection
        Dim LoginCompleteParametersArray As String()
    End Structure

    Class WebRequestHandler

        Structure WebRequestStruct
            Dim webReq As HttpWebRequest
            Dim webResp As HttpWebResponse
            Dim SStream As Stream
            Dim SReader As StreamReader
        End Structure

        Structure WebResponseStruct
            Dim innerHTML As String
            Dim RespCookie As CookieCollection
        End Structure

        Private webRR As WebRequestStruct

        Public Sub New(ByVal u As String, ByVal c As CookieCollection)
            webRR.webReq = HttpWebRequest.Create(u)
            With webRR.webReq
                .CookieContainer = New CookieContainer()
                .CookieContainer.Add(c)
                .ContentType = "application/x-www-form-urlencoded"
                .Headers.Add("Accept-Language", "zh-cn")
                .Method = "GET"
            End With
        End Sub
        Public Sub New(ByVal u As String, ByVal c As String)
            webRR.webReq = HttpWebRequest.Create(u)
            Functions.SetCookieHeaders(webRR.webReq, c)
            With webRR.webReq
                .ContentType = "application/x-www-form-urlencoded"
                .Headers.Add("Accept-Language", "zh-cn")
                .Method = "GET"
            End With
        End Sub
        Public Function GetSrcOnly() As String

        End Function
        Public Function GetSrcAndCookie() As WebResponseStruct
            webRR.webResp = webRR.webReq.GetResponse()
            webRR.SStream = webRR.webResp.GetResponseStream()
            Dim nWRS As WebResponseStruct
            nWRS.innerHTML = New StreamReader(webRR.SStream, Encoding.GetEncoding("UTF-8")).ReadToEnd
            nWRS.RespCookie = webRR.webResp.Cookies
            Return nWRS
        End Function
    End Class

#End Region

End Module
