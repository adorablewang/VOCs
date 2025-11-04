Imports System.Text

''' <summary>
''' 用于国标212协议的通讯
''' </summary>
Public Class GB212Protocol
    ''' <summary>
    ''' ST编码（系统编号）
    ''' </summary>
    Private STNumber As String = "31"

    ''' <summary>
    ''' 发送至服务器密码
    ''' </summary>
    Private PWord As String = "123456"

    ''' <summary>
    ''' MN号，设备唯一标识，EPC-96编码，总共由24个字符组成（0-9，A-F)
    ''' </summary>
    Private MNNumber As String = "20230212049_1"

    ''' <summary>
    ''' 用户名
    ''' </summary>
    Private _userName As String = "湖北方圆VOC在线监测"

    Property UserName As String
        Get
            Return _userName
        End Get
        Set(value As String)
            _userName = value
        End Set
    End Property


    ''' <summary>
    ''' 212协议中使用的用户密码
    ''' </summary>
    ''' <returns></returns>
    Public Property UserPWord As String
        Get
            Return PWord
        End Get
        Set(value As String)
            PWord = value
        End Set
    End Property

    ''' <summary>
    ''' 212协议中用户MN号
    ''' </summary>
    ''' <returns></returns>
    Public Property UserMN As String
        Get
            Return MNNumber
        End Get
        Set(value As String)
            MNNumber = value
        End Set
    End Property

    ''' <summary>
    ''' 用户类型，根据212协议中进行修改
    ''' </summary>
    ''' <returns></returns>
    Public Property UserType As String
        Get
            Return STNumber
        End Get
        Set(value As String)
            STNumber = value
        End Set
    End Property
    ''' <summary>
    ''' 212协议中的校验码，
    ''' </summary>
    ''' <param name="strData">表示需要校验的字符串</param>
    ''' <param name="dataLen">表示需要校验的字符长度</param>
    ''' <returns>校验码</returns>
    Public Function CRC16To212(ByVal strData As String, ByVal dataLen As Integer) As String
        Dim bt() As Byte = Encoding.ASCII.GetBytes(strData)
        Dim crc As UShort = &HFFFF
        Dim poly As UShort = &HA001
        Dim iIndex As Integer
        Dim iIndex2 As Integer

        For iIndex = 0 To dataLen - 1
            crc >>= 8
            crc = crc Xor bt(iIndex)
            For iIndex2 = 0 To 7
                If (crc And &H1) = &H1 Then
                    crc = (crc >> 1) Xor poly
                Else
                    crc >>= 1
                End If
            Next
        Next
        Return crc.ToString("X4")
    End Function
    ''' <summary>
    ''' 发送实时数据，dataSegment表示数据段的数据
    ''' </summary>
    ''' <param name="dataSegment">实时数据段字符串</param>
    ''' <returns></returns>
    Public Function RealTimeData_2011(ByVal dataSegment As String) As String
        Dim strbuff As New StringBuilder
        Dim crcbuff As String

        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=" & STNumber & ";")
        strbuff.Append("CN=2011;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=5;")
        strbuff.Append("CP=&&DataTime=" & dataSegment & "&&")
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr & vbLf)
        Return strbuff.ToString()
    End Function

    ''' <summary>
    ''' 发送分钟数据到服务器
    ''' </summary>
    ''' <param name="dataSegment">分钟数据段字符串</param>
    ''' <returns></returns>
    Public Function MinuteData_2051(ByVal dataSegment As String) As String
        Dim strbuff As New StringBuilder
        Dim crcbuff As String

        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=" & STNumber & ";")
        strbuff.Append("CN=2051;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=5;")
        strbuff.Append("CP=&&DataTime=" & dataSegment & "&&")
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        '包结构中的“数据段长度 4个ASCII码字符”，比如长255，则写入“0255”4个字符
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        '这是包头，固定值“##”
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        '这是包尾
        'ASCII码"\r"回车
        strbuff.Append(vbCr)
        'ASCII码"\n"换行
        strbuff.Append(vbLf)
        Return strbuff.ToString()
    End Function
    ''' <summary>
    ''' 发送小时数据到服务器
    ''' HJ212-2017协议标准报文
    ''' 请求编码QN=yyyyMMddHHmmssfff,长度20用来唯一标识一次命令交互
    ''' 系统编码ST，长度5 表示系统编号
    ''' 命令编号CN，长度7
    ''' 访问密码PW，长度9
    ''' 设备唯一标识MN,长度27 表示监测点编号
    ''' </summary>
    ''' <param name="dataSegment">小时数据段字符串</param>
    ''' <returns></returns>
    Public Function HourData_2061(ByVal dataSegment As String) As String
        Dim strbuff As New StringBuilder
        Dim crcbuff As String

        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=" & STNumber & ";")
        strbuff.Append("CN=2061;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=5;")
        strbuff.Append("CP=&&DataTime=" & dataSegment & "&&")
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        Return strbuff.ToString()
    End Function

    ''' <summary>
    ''' 发送日数据到服务器
    ''' </summary>
    ''' <param name="dataSegment">日数据段字符串</param>
    ''' <returns></returns>
    Public Function DailyData_2031(ByVal dataSegment As String) As String
        Dim strbuff As New StringBuilder
        Dim crcbuff As String

        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=" & STNumber & ";")
        strbuff.Append("CN=2031;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=5;")
        strbuff.Append("CP=&&DataTime=" & dataSegment & "&&")

        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        Return strbuff.ToString()
    End Function

    ''' <summary>
    ''' 返回现场工控机的系统时间
    ''' </summary>
    ''' <returns>返回字符串数组，包含三组需要发送的数据</returns>
    Public Function GetStationDateTime() As String()
        Dim str(2) As String
        Dim strbuff As New StringBuilder
        Dim crcbuff As String

        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=91;")
        strbuff.Append("CN=9011;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=4;")
        strbuff.Append("CP=&&QnRtn=1&&")
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        str(0) = strbuff.ToString()

        strbuff.Clear()
        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=32;")
        strbuff.Append("CN=1011;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=4;")
        strbuff.Append("CP=&&SystemTime=" & DateTime.Now.ToString("yyyyMMddHHmmss") & "&&")
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        str(1) = strbuff.ToString()

        strbuff.Clear()
        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=91;")
        strbuff.Append("CN=9012;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=4;")
        strbuff.Append("CP=&&ExeRtn=1&&")
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        str(2) = strbuff.ToString()

        Return str
    End Function

    ''' <summary>
    ''' 设置现场工控机中的系统时间，返回字符串数组
    ''' </summary>
    ''' <returns></returns>
    Public Function SetStationDateTime() As String()
        Dim str(1) As String
        Dim strbuff As New StringBuilder
        Dim crcbuff As String

        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=91;")
        strbuff.Append("CN=9011;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=4;")
        strbuff.Append("CP=&&QnRtn=1&&")
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        str(0) = strbuff.ToString()

        strbuff.Clear()
        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=91;")
        strbuff.Append("CN=9012;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=4;")
        strbuff.Append("CP=&&ExeRtn=1&&")
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        str(1) = strbuff.ToString()

        Return str
    End Function

    ''' <summary>
    ''' 提取现场公共机实时数据上传时间间隔
    ''' </summary>
    ''' <returns></returns>
    Public Function GetRTDataSendTime() As String()
        Dim str(2) As String
        Dim strbuff As New StringBuilder
        Dim crcbuff As String

        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=91;")
        strbuff.Append("CN=9011;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=4;")
        strbuff.Append("CP=&&QnRtn=1&&")
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        str(0) = strbuff.ToString()

        strbuff.Clear()
        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=32;")
        strbuff.Append("CN=1061;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=4;")
        strbuff.Append("CP=&&RtdInterval=30&&")    '此处需要一个变量，暂时没有处理
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        str(1) = strbuff.ToString()

        strbuff.Clear()
        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=91;")
        strbuff.Append("CN=9012;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=4;")
        strbuff.Append("CP=&&ExeRtn=1&&")
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        str(2) = strbuff.ToString()

        Return str
    End Function

    ''' <summary>
    ''' 设置现场工控机上传实时数据间隔时间
    ''' </summary>
    ''' <returns></returns>
    Public Function SetRTDataSendTime() As String()
        Dim str(1) As String
        Dim strbuff As New StringBuilder
        Dim crcbuff As String

        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=91;")
        strbuff.Append("CN=9011;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=4;")
        strbuff.Append("CP=&&QnRtn=1&&")
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        str(0) = strbuff.ToString()

        strbuff.Clear()
        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=91;")
        strbuff.Append("CN=9012;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=4;")
        strbuff.Append("CP=&&ExeRtn=1&&")
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        str(1) = strbuff.ToString()

        Return str
    End Function

    Public Function SetPassWord() As String()
        Dim str(1) As String
        Dim strbuff As New StringBuilder
        Dim crcbuff As String

        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=91;")
        strbuff.Append("CN=9011;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=4;")
        strbuff.Append("CP=&&QnRtn=1&&")
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        str(0) = strbuff.ToString()

        strbuff.Clear()
        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=91;")
        strbuff.Append("CN=9012;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=4;")
        strbuff.Append("CP=&&ExeRtn=1&&")
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        str(1) = strbuff.ToString()

        Return str
    End Function

    ''' <summary>
    ''' 提取现场站点的唯一识别码MN号
    ''' </summary>
    ''' <returns></returns>
    Public Function GetStationMN() As String()
        Dim str(2) As String
        Dim strbuff As New StringBuilder
        Dim crcbuff As String

        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=91;")
        strbuff.Append("CN=9011;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=4;")
        strbuff.Append("CP=&&QnRtn=1&&")
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        str(0) = strbuff.ToString()

        strbuff.Clear()
        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=32;")
        strbuff.Append("CN=3019;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=4;")
        strbuff.Append("CP=&&RtdInterval=" & MNNumber & "&&")    '此处需要一个变量，暂时没有处理
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        str(1) = strbuff.ToString()

        strbuff.Clear()
        strbuff.Append("QN=" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuff.Append("ST=91;")
        strbuff.Append("CN=9012;")
        strbuff.Append("PW=" & PWord & ";")
        strbuff.Append("MN=" & MNNumber & ";")
        strbuff.Append("Flag=4;")
        strbuff.Append("CP=&&ExeRtn=1&&")
        crcbuff = CRC16To212(strbuff.ToString(), strbuff.Length)
        strbuff.Insert(0, strbuff.Length.ToString("D4"))
        strbuff.Insert(0, "##")
        strbuff.Append(crcbuff)
        strbuff.Append(vbCr)
        strbuff.Append(vbLf)
        str(2) = strbuff.ToString()

        Return str
    End Function
End Class

