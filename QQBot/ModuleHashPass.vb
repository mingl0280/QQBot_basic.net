Imports System.Text
Imports System.IO

Module ModuleHashPass

    ''' <summary>
    ''' 一次md5加密
    ''' </summary>
    ''' <param name="md5_str">需要加密的文本</param>
    ''' <returns></returns>
    Public Function MD5_Encrypt(md5_str As String) As String
        Dim md5 As System.Security.Cryptography.MD5 = System.Security.Cryptography.MD5CryptoServiceProvider.Create()
        Dim bytes As Byte() = System.Text.Encoding.ASCII.GetBytes(md5_str)
        Dim bytes1 As Byte() = md5.ComputeHash(bytes)

        Dim stringBuilder As System.Text.StringBuilder = New StringBuilder()
        For Each item In bytes1
            stringBuilder.Append(item.ToString("x").PadLeft(2, "0"c))
        Next
        Return stringBuilder.ToString().ToUpper()
    End Function
    ''' <summary>
    ''' 将字节流加密
    ''' </summary>
    ''' <param name="md5_bytes">需要加密的字节流</param>
    ''' <returns></returns>
    Public Function MD5_Encrypt(md5_bytes As Byte()) As String
        Dim md5 As System.Security.Cryptography.MD5 = System.Security.Cryptography.MD5CryptoServiceProvider.Create()

        Dim bytes1 As Byte() = md5.ComputeHash(md5_bytes)
        Dim stringBuilder As System.Text.StringBuilder = New StringBuilder()
        For Each item In bytes1
            stringBuilder.Append(item.ToString("x").PadLeft(2, "0"c))
        Next
        Return stringBuilder.ToString().ToUpper()

    End Function
    ''' <summary>
    ''' 获取文本的md5字节流
    ''' </summary>
    ''' <param name="md5_str">需要加密成Md5d的文本</param>
    ''' <returns></returns>
    Private Function MD5_GetBytes(md5_str As String) As Byte()
        Dim md5 As System.Security.Cryptography.MD5 = System.Security.Cryptography.MD5CryptoServiceProvider.Create()
        Dim bytes As Byte() = System.Text.Encoding.ASCII.GetBytes(md5_str)
        Return md5.ComputeHash(bytes)


    End Function
    ''' <summary>
    ''' 加密成md5字节流之后转换成文本
    ''' </summary>
    ''' <param name="md5_str"></param>
    ''' <returns></returns>
    Private Function Encrypt_1(md5_str As String) As String
        Dim md5 As System.Security.Cryptography.MD5 = System.Security.Cryptography.MD5CryptoServiceProvider.Create()
        Dim bytes As Byte() = System.Text.Encoding.ASCII.GetBytes(md5_str)
        bytes = md5.ComputeHash(bytes)
        Dim stringBuilder As System.Text.StringBuilder = New StringBuilder()
        For Each item In bytes
            stringBuilder.Append("\x")
            stringBuilder.Append(item.ToString("x2"))
        Next
        Return stringBuilder.ToString()
    End Function
    Public Class ByteBuffer
        Private _buffer As Byte()
        ''' <summary>
        ''' 获取同后备存储区连接的基础流
        ''' </summary>
        Public BaseStream As Stream

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        Public Sub New()
            Me.BaseStream = New MemoryStream()
            Me._buffer = New Byte(15) {}
        End Sub

        ''' <summary>
        ''' 设置当前流中的位置
        ''' </summary>
        ''' <param name="offset">相对于origin参数字节偏移量</param>
        ''' <param name="origin">System.IO.SeekOrigin类型值,指示用于获取新位置的参考点</param>
        ''' <returns></returns>
        Public Overridable Function Seek(offset As Integer, origin As SeekOrigin) As Long
            Return Me.BaseStream.Seek(CLng(offset), origin)
        End Function

        ''' <summary>
        ''' 检测是否还有可用字节
        ''' </summary>
        ''' <returns></returns>
        Public Function Peek() As Boolean
            Return If(BaseStream.Position >= BaseStream.Length, False, True)
        End Function

        ''' <summary>
        ''' 将整个流内容写入字节数组，而与 Position 属性无关。
        ''' </summary>
        ''' <returns></returns>
        Public Function ToByteArray() As Byte()
            Dim org As Long = BaseStream.Position
            BaseStream.Position = 0
            Dim ret As Byte() = New Byte(BaseStream.Length - 1) {}
            BaseStream.Read(ret, 0, ret.Length)
            BaseStream.Position = org
            Return ret
        End Function

#Region "写流方法"
        ''' <summary>
        ''' 压入一个布尔值,并将流中当前位置提升1
        ''' </summary>
        ''' <param name="value"></param>
        Public Sub Put(value As Boolean)
            Me._buffer(0) = If(value, CByte(1), CByte(0))
            Me.BaseStream.Write(_buffer, 0, 1)
        End Sub

        ''' <summary>
        ''' 压入一个Byte,并将流中当前位置提升1
        ''' </summary>
        ''' <param name="value"></param>
        Public Sub Put(value As [Byte])
            Me.BaseStream.WriteByte(value)
        End Sub
        ''' <summary>
        ''' 压入Byte数组,并将流中当前位置提升数组长度
        ''' </summary>
        ''' <param name="value">字节数组</param>
        Public Sub Put(value As [Byte]())
            If value Is Nothing Then
                Throw New ArgumentNullException("value")
            End If
            Me.BaseStream.Write(value, 0, value.Length)
        End Sub
        ''' <summary>
        ''' Puts the int.
        ''' </summary>
        ''' <param name="value">The value.</param>
        Public Sub PutInt(value As Integer)
            PutInt(CUInt(value))
        End Sub
        ''' <summary>
        ''' 压入一个int,并将流中当前位置提升4
        ''' </summary>
        ''' <param name="value"></param>
        Public Sub PutInt(value As UInteger)
            Me._buffer(0) = CByte(value >> &H18)
            Me._buffer(1) = CByte(value >> 16)
            Me._buffer(2) = CByte(value >> 8)
            Me._buffer(3) = CByte(value)
            Me.BaseStream.Write(Me._buffer, 0, 4)
        End Sub
        ''' <summary>
        ''' Puts the int.
        ''' </summary>
        ''' <param name="index">The index.</param>
        ''' <param name="value">The value.</param>
        Public Sub PutInt(index As Integer, value As UInteger)
            Dim pos As Integer = CInt(Me.BaseStream.Position)
            Seek(index, SeekOrigin.Begin)
            PutInt(value)
            Seek(pos, SeekOrigin.Begin)
        End Sub

#End Region

#Region "读流方法"

        ''' <summary>
        ''' 读取Byte值,并将流中当前位置提升1
        ''' </summary>
        ''' <returns></returns>
        Public Function [Get]() As Byte
            Return CByte(BaseStream.ReadByte())
        End Function

#End Region
    End Class

    Public Function MD5_QQ_2_Encrypt(uin As Integer, password As String, verifyCode As String) As String

        Dim buffer As New ByteBuffer()
        buffer.Put(MD5_GetBytes(password))
        'buffer.Put(Encoding.UTF8.GetBytes(password));
        buffer.PutInt(0)
        buffer.PutInt(uin)
        Dim bytes As Byte() = buffer.ToByteArray()
        Dim md5_1 As String = MD5_Encrypt(bytes)
        '将混合后的字节流进行一次md5加密
        Dim result As String = MD5_Encrypt(md5_1 & verifyCode.ToUpper())
        '再用加密后的结果与大写的验证码一起加密一次
        Return result

    End Function


End Module
