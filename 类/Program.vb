Imports Ini.Net
Imports System.IO
Imports System.IO.Ports


Module program

    '配置文件的路径
    Public ReadOnly IniPathName As String = Application.StartupPath & "\config.ini"
    Public ReadOnly ConfigFilePath As String = Application.StartupPath & "\sqlite.db"

    '选择与服务器的通讯协议 为0时候，实验室认证，为1时候，现场认证
    Public SelectProtocol As Integer

    '显示所有的参数 实时参数
    Public saveparameter As New VOCParameterName
    '为计算每5秒钟的分钟平均值
    Public minuteAvg As New List(Of VOCParameterName)
    '关键参数设置
    Public keyParameter As New VOCKeyParameter
    '量程限值设置
    Public rangeParameter As New VOCRangeSetting
    '参数报警设置
    Public alarmParameter As New AlarmSetting
    '烟气反吹设置
    Public purgeSetting As New PurgeGasSetting
    '校准参数设置
    Public calibrationPara As New CalibrationSetting
    '212协议中的用户信息
    Public senddatatoserver As New GB212Protocol

    Public analyzerdata As New Analyzer '分析仪采集 甲烷（标杆)  ，总烃(标杆），非甲烷总烃

    Public particulatedata As New ParticulateDevice  '颗粒物浓度获取设备

    'Public analyzerFY As New ClassFYCH4
    Public plcdata As New PLCCommu
    Public o2data As New O2COMMU  '采集氧气，有的氧气仪还可以采集湿度
    Public datasamplinganalyzer As New SerialPort '数采仪串口
    Public plctcp As PLC_TCP '控制整个设备上的一些灯

    Public maindisplay As Func(Of String, String)

    '用户权限管理，保存用户类型(管理员或者巡视员)
    '默认是管理员登录
    Public UserRole As EnumDefine.UserRoles = EnumDefine.UserRoles.Administrator



    Sub Main()
        Dim createdNew As Boolean
        ' 创建mutex
        'Dim mutex As New System.Threading.Mutex(True, "污染源VOCs", createdNew)

        Dim mutex As New System.Threading.Mutex(True, Application.ProductName, createdNew)

        If createdNew = False Then
            Return
        End If

        IniDb()
        'isDirExist(Application.StartupPath & "\Log")
        isDirExist(Application.StartupPath & "\Log\communication")
        isDirExist(Application.StartupPath & "\Log\alarm")
        isDirExist(Application.StartupPath & "\Log\Operation")
        isDirExist(Application.StartupPath & "\Log\power")
        Application.Run(New FormMain)
        ' 释放mutex
        mutex.ReleaseMutex()
    End Sub

    Public Function GetRoundedRectPath(ByVal rect As Rectangle, ByVal radius As Integer) As System.Drawing.Drawing2D.GraphicsPath
        rect.Offset(-1, -1)
        Dim RoundRect As New Rectangle(rect.Location, New Size(radius - 1, radius - 1))
        Dim path As New System.Drawing.Drawing2D.GraphicsPath

        path.AddArc(RoundRect, 180, 90)     '左上角

        RoundRect.X = rect.Right - radius   '右上角
        path.AddArc(RoundRect, 270, 90)

        RoundRect.Y = rect.Bottom - radius  '右下角
        path.AddArc(RoundRect, 0, 90)

        RoundRect.X = rect.Left             '左下角
        path.AddArc(RoundRect, 90, 90)

        path.CloseFigure()
        Return path
    End Function

    ''' <summary>
    ''' 用来控制控件圆角矩形，offset为了改变坐标
    ''' </summary>
    ''' <param name="rect"></param>
    ''' <param name="radius"></param>
    ''' <returns></returns>
    Public Function GetRoundedRectPath(ByVal rect As Rectangle, ByVal radius As Integer, cut As Boolean) As System.Drawing.Drawing2D.GraphicsPath
        '如果需要圆角矩形控件
        If cut Then rect.Offset(-1, -1)

        Dim RoundRect As New Rectangle(rect.Location, New Size(radius - 1, radius - 1))
        Dim path As New System.Drawing.Drawing2D.GraphicsPath

        path.AddArc(RoundRect, 180, 90)     '左上角

        RoundRect.X = rect.Right - radius   '右上角
        path.AddArc(RoundRect, 270, 90)

        RoundRect.Y = rect.Bottom - radius  '右下角
        path.AddArc(RoundRect, 0, 90)

        RoundRect.X = rect.Left             '左下角
        path.AddArc(RoundRect, 90, 90)

        path.CloseFigure()

        Return path
    End Function

    Public Function CreateRoundedRectPath(rect As Rectangle, radius As Integer, lu As Boolean, ru As Boolean, rd As Boolean, ld As Boolean) As System.Drawing.Drawing2D.GraphicsPath
        Dim RoundRect As New System.Drawing.Drawing2D.GraphicsPath

        '上端和右上角
        If ru Then
            RoundRect.AddLine(rect.Left + radius - 2, rect.Top - 1, rect.Right - radius, rect.Top - 1)          '顶端 
            RoundRect.AddArc(rect.Right - radius, rect.Top - 1, radius, radius, 270, 90)                        '右上角 
        Else
            RoundRect.AddLine(rect.Left + radius - 2, rect.Top - 1, rect.Right, rect.Top - 1)          '顶端
            RoundRect.AddLine(rect.Right, rect.Top, rect.Right, rect.Top + radius)
        End If

        If rd Then
            RoundRect.AddLine(rect.Right, rect.Top + radius, rect.Right, rect.Bottom - radius)                  '右边 
            RoundRect.AddArc(rect.Right - radius, rect.Bottom - radius, radius, radius, 0, 90)                  '右下角
        Else
            RoundRect.AddLine(rect.Right, rect.Top + radius, rect.Right, rect.Bottom - radius)                  '右边 
            RoundRect.AddLine(rect.Right, rect.Bottom, rect.Right - radius, rect.Bottom)                '底边 
        End If

        If ld Then
            RoundRect.AddLine(rect.Right - radius, rect.Bottom, rect.Left + radius, rect.Bottom)                '底边 
            RoundRect.AddArc(rect.Left - 1, rect.Bottom - radius, radius, radius, 90, 90)                       '左下角
        Else
            RoundRect.AddLine(rect.Right - radius, rect.Bottom, rect.Left + radius, rect.Bottom)                '底边 
            RoundRect.AddLine(rect.Left - 1, rect.Bottom, rect.Left - 1, rect.Bottom - radius)
        End If

        If lu Then
            RoundRect.AddLine(rect.Left - 1, rect.Top + radius, rect.Left - 1, rect.Bottom - radius)            '左边 
            RoundRect.AddArc(rect.Left - 1, rect.Top - 1, radius, radius, 180, 90)                              '左上角 
        Else
            RoundRect.AddLine(rect.Left - 1, rect.Top + radius, rect.Left - 1, rect.Bottom - radius)
            RoundRect.AddLine(rect.Left - 1, rect.Top - 1, rect.Left + radius, rect.Top - 1)
        End If

        Return RoundRect
    End Function

#Region "读取配置文件"
    ''' <summary>
    ''' 初始化INI文件中的所有项，
    ''' </summary>
    ''' <param name="section"></param>
    ''' <param name="key"></param>
    ''' <param name="value">如果参数不存在，默认的参数</param>
    ''' <returns></returns>
    Public Function ReadConfigIni(section As String, key As String, value As String) As String
        Dim InitializeName As New IniFile(IniPathName)
        Dim strbuf As String

        strbuf = InitializeName.ReadString(section, key)
        If "" = strbuf Then
            InitializeName.WriteString(section, key, value)
            Return value
        Else
            Return strbuf
        End If
    End Function
#End Region

    ''' <summary>
    ''' 初始化系统配置文件路径为  ConfigFilePath
    ''' </summary>
    Private Sub IniDb()

        Dim sqlHelper As SQLiteHelper

        sqlHelper = New SQLiteHelper(ConfigFilePath)
        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        Dim sqlStrs As New ArrayList

        sqlStrs.Add("CREATE TABLE IF NOT EXISTS UserInfo (
	        id INTEGER PRIMARY KEY AUTOINCREMENT,
            UserName TEXT,
            UserRight TEXT,
            UserPW TEXT,
	        AddTime DATETIME);")

        sqlStrs.Add("CREATE TABLE IF NOT EXISTS 分钟数据 (
	        id INTEGER PRIMARY KEY AUTOINCREMENT,
	        samplingtime TEXT,
            ch4conc_h REAL,
            ch4conc REAL,
            ch4conc_z REAL,
            ch4conc_p REAL,
            nmhcconc_h REAL,
            nmhcconc REAL,
            nmhcconc_z REAL,
            nmhcconc_p REAL,
            tchconc_h REAL,
            tchconc REAL,
            tchconc_z REAL,
            tchconc_p REAL,        
            particulateconc_h REAL,
            particulateconc REAL,
            particulateconc_z REAL,
            particulateconc_p REAL,
            o2conc_h REAL,
            o2conc REAL,
            tempconc REAL,
            pressconc REAL,
            speedconc REAL,
            flowconc_g REAL,
            flowconc_b REAL,
            humiconc REAL,
            fuheconc REAL,
            REMARK TEXT);")

        sqlStrs.Add("CREATE TABLE IF NOT EXISTS 小时数据 (
	        id INTEGER PRIMARY KEY AUTOINCREMENT,
	        samplingtime TEXT,
            ch4conc_h REAL,
            ch4conc REAL,
            ch4conc_z REAL,
            ch4conc_p REAL,
            nmhcconc_h REAL,
            nmhcconc REAL,
            nmhcconc_z REAL,
            nmhcconc_p REAL,
            tchconc_h REAL,
            tchconc REAL,
            tchconc_z REAL,
            tchconc_p REAL,
            particulateconc_h REAL,
            particulateconc REAL,
            particulateconc_z REAL,
            particulateconc_p REAL,
            o2conc_h REAL,
            o2conc REAL,
            tempconc REAL,
            pressconc REAL,
            speedconc REAL,
            flowconc_g REAL,
            flowconc_b REAL,
            humiconc REAL,
            fuheconc REAL,
            REMARK TEXT);")

        sqlStrs.Add("CREATE TABLE IF NOT EXISTS 日数据 (
	        id INTEGER PRIMARY KEY AUTOINCREMENT,
	        samplingtime TEXT,
            ch4conc_h REAL,
            ch4conc REAL,
            ch4conc_z REAL,
            ch4conc_p REAL,
            nmhcconc_h REAL,
            nmhcconc REAL,
            nmhcconc_z REAL,
            nmhcconc_p REAL,
            tchconc_h REAL,
            tchconc REAL,
            tchconc_z REAL,
            tchconc_p REAL,
            particulateconc_h REAL,
            particulateconc REAL,
            particulateconc_z REAL,
            particulateconc_p REAL,
            o2conc_h REAL,
            o2conc REAL,
            tempconc REAL,
            pressconc REAL,
            speedconc REAL,
            flowconc_g REAL,
            flowconc_b REAL,
            humiconc REAL,
            fuheconc REAL,
            REMARK TEXT);")

        sqlStrs.Add("CREATE TABLE IF NOT EXISTS 月数据 (
	        id INTEGER PRIMARY KEY AUTOINCREMENT,
	        samplingtime TEXT,
            ch4conc_h REAL,
            ch4conc REAL,
            ch4conc_z REAL,
            ch4conc_p REAL,
            nmhcconc_h REAL,
            nmhcconc REAL,
            nmhcconc_z REAL,
            nmhcconc_p REAL,
            tchconc_h REAL,
            tchconc REAL,
            tchconc_z REAL,
            tchconc_p REAL,
            particulateconc_h REAL,
            particulateconc REAL,
            particulateconc_z REAL,
            particulateconc_p REAL,
            o2conc_h REAL,
            o2conc REAL,
            tempconc REAL,
            pressconc REAL,
            speedconc REAL,
            flowconc_g REAL,
            flowconc_b REAL,
            humiconc REAL,
            fuheconc REAL,
            REMARK TEXT);")

        sqlStrs.Add("CREATE TABLE IF NOT EXISTS 关键参数记录 (
         id INTEGER PRIMARY KEY AUTOINCREMENT,
            samplingtime TEXT,
            SamplingArea REAL,
            Press REAL,
            CoeffO2  REAL,
            SPEED REAL,
            Pitol REAL,
            LimitO2 REAL,
            ConvertFlag REAL,
            REMARK TEXT);")

        sqlStrs.Add("CREATE TABLE IF NOT EXISTS 量程记录 (
	        id INTEGER PRIMARY KEY AUTOINCREMENT,
	        samplingtime TEXT,
            tchlow REAL,
            tchup REAL,
            o2low REAL,
            o2up REAL,
            templow REAL,
            tempup REAL,
            presslow REAL,
            pressup REAL,
            speedlow REAL,
            speedup REAL,
            humilow REAL,
            humiup REAL);")

        sqlStrs.Add("CREATE TABLE IF NOT EXISTS 事件记录 (
         id INTEGER PRIMARY KEY AUTOINCREMENT,
            samplingtime TEXT,
            username TEXT,
            info TEXT,
            infotype INTEGER);")

        sqlStrs.Add("CREATE TABLE IF NOT EXISTS 报警记录 (
         id INTEGER PRIMARY KEY AUTOINCREMENT,
            samplingtime TEXT,
            username TEXT,
            info TEXT,
            infotype INTEGER);")

        sqlStrs.Add("CREATE TABLE IF NOT EXISTS 量程修改 (
         id INTEGER PRIMARY KEY AUTOINCREMENT,
            samplingtime TEXT,
            username TEXT,
            paraname TEXT,
            rangelow REAL,
            rangeup REAL，
            alarmlow REAL,
            alarmup REAL);")

        'sqlStrs.Add("CREATE TABLE IF NOT EXISTS 校准参数 (
        ' id INTEGER PRIMARY KEY AUTOINCREMENT,
        ' ParaNum INTEGER,
        '    ParaName TEXT,
        '    CalibrationConc REAL,
        '    PeakArea REAL,
        ' REMARK DATETIME);")
        ''这是一个测试的数据表
        'sqlStrs.Add("CREATE TABLE IF NOT EXISTS haha (
        ' id INTEGER,
        ' ceshi TEXT,
        ' kaishi TEXT);")
        sqlHelper.IUD(sqlStrs)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If

        sqlHelper.Close()
    End Sub

    Public Sub Export_csv(dgv As DGV_MulCapRowMerge, s1 As Stream)
        Using sw1 As New StreamWriter(s1, System.Text.Encoding.Default)

            Dim columntitle As String = ""

            Try
                For i As Integer = 0 To dgv.Columns.Count - 1
                    If i > 0 Then
                        columntitle += ","
                    End If
                    columntitle += dgv.Columns(i).HeaderText
                Next

                columntitle.Remove(columntitle.Length - 1)
                sw1.WriteLine(columntitle)

                For j As Integer = 0 To dgv.Rows.Count - 1
                    Dim columevalue As String = ""

                    For k As Integer = 0 To dgv.Columns.Count - 1
                        If k > 0 Then
                            columevalue += ","
                        End If

                        If dgv.Rows(j).Cells(k).Value Is Nothing Then
                            columevalue += ""
                        Else
                            Dim m As String = dgv.Rows(j).Cells(k).Value.ToString().Trim()
                            columevalue += m.Replace(",", ",")
                        End If
                    Next

                    columevalue.Remove(columevalue.Length - 1)
                    sw1.WriteLine(columevalue)
                Next
                sw1.Close()
                s1.Close()

            Catch ex As Exception

            Finally
                sw1.Close()
                s1.Close()
            End Try
        End Using
    End Sub
    Public Function strChar(ByVal Str As String) As String
        Dim ReStr As String = Str
        If Not ((Str = "") OrElse (Str Is Nothing)) Then
            If (Str.Length > 0) Then

                ReStr = Str.Replace("" & vbLf, "").Replace("" & vbTab, "").Replace("" & vbCr, "")
            End If
        End If

        Return ReStr
    End Function

    Public Sub Export_csv(dgv As DataGridView, s1 As Stream)
        Using sw1 As New StreamWriter(s1, System.Text.Encoding.Default)

            Dim columntitle As String = ""

            Try
                For i As Integer = 0 To dgv.Columns.Count - 1
                    If i > 0 Then
                        columntitle += ","
                    End If
                    Dim fileName = strChar(dgv.Columns(i).HeaderText)
                    columntitle += fileName
                Next

                columntitle.Remove(columntitle.Length - 1)
                sw1.WriteLine(columntitle)

                For j As Integer = 0 To dgv.Rows.Count - 1
                    Dim columevalue As String = ""

                    For k As Integer = 0 To dgv.Columns.Count - 1
                        If k > 0 Then
                            columevalue += ","
                        End If

                        If dgv.Rows(j).Cells(k).Value Is Nothing Then
                            columevalue += ""
                        Else
                            Dim m As String = dgv.Rows(j).Cells(k).Value.ToString().Trim()
                            columevalue += m.Replace(",", ",")
                        End If
                    Next

                    columevalue.Remove(columevalue.Length - 1)
                    sw1.WriteLine(columevalue)
                Next
                sw1.Close()
                s1.Close()

            Catch ex As Exception

            Finally
                sw1.Close()
                s1.Close()
            End Try
        End Using
    End Sub


    Public UnicodeSuperscriptConversion As New Dictionary(Of String, String) From
{
    {"0", ChrW(&H2070)},
    {"1", ChrW(&HB9)},
    {"2", ChrW(&HB2)},
    {"3", ChrW(&HB3)},
    {"4", ChrW(&H2074)},
    {"5", ChrW(&H2075)},
    {"6", ChrW(&H2076)},
    {"7", ChrW(&H2077)},
    {"8", ChrW(&H2078)},
    {"9", ChrW(&H2079)},
    {".", ChrW(&HB7)},
    {"-", ChrW(&H207B)}
}

    Public UnicodeLimitConversion As New Dictionary(Of String, String) From
        {
        {"0", "₀"},
        {"1", "₁"},
        {"2", "₂"},
        {"3", "₃"},
        {"4", "₄"},
        {"5", "₅"},
        {"6", "₆"},
        {"7", "₇"},
        {"8", "₈"},
        {"9", "₉"}
    }

    '创建文件夹，如果不存在创建
    Public Sub isDirExist(ByVal strPath As String)
        Dim strDirTemp As String()
        strDirTemp = strPath.Split("\")
        strPath = String.Empty
        For i As Integer = 0 To strDirTemp.Length - 1
            ' 判断数组内容.目的是防止输入的strPath内容如:c:\abc\123\ 最后一位也是"\"
            If strDirTemp(i) <> "" Then
                strPath += strDirTemp(i) & "\"
            End If
        Next

        ' 判断文件夹是否存在
        If Not Directory.Exists(strPath) Then
            Directory.CreateDirectory(strPath)
        End If
    End Sub

    ''' <summary>
    ''' 写日志到文本文件中
    ''' </summary>
    ''' <param name="fileName">文件名</param>
    ''' <param name="fileType">文件类型</param>
    ''' <param name="info">文件内容</param>
    Public Sub WriteLogToText(fileName As String, fileType As String, info As String, Optional append As Boolean = True)

        Dim filepath As String = Application.StartupPath & "Log\"
        If String.Equals(fileType.ToLower, "communication") Then
            filepath = Application.StartupPath & "\Log\communication\" & fileName & ".TXT"
        End If

        If String.Equals(fileType.ToLower, "alarm") Then
            filepath = Application.StartupPath & "\Log\alarm\" & fileName & ".TXT"
        End If

        If String.Equals(fileType.ToLower, "operation") Then
            filepath = Application.StartupPath & "\Log\operation\" & fileName & ".TXT"
        End If

        If String.Equals(fileType.ToLower, "power") Then
            filepath = Application.StartupPath & "\Log\power\" & fileName & ".TXT"
        End If

        Dim sw As StreamWriter = New StreamWriter(filepath, append)
        Dim strbuf As String
        strbuf = Now.ToString("HH:mm:ss") & "  " & info
        sw.WriteLine(strbuf)
        sw.Flush()
        sw.Close()
        sw = Nothing
    End Sub


    Public Class ModbusConvert
        ''' <summary>
        ''' 浮点数变为ushort()
        ''' </summary>
        ''' <param name="data">浮点数</param>
        ''' <returns></returns>
        Public Function FloatToUshort(data As Single) As UShort()

            Dim u(1) As UShort

            Dim b As Byte() = BitConverter.GetBytes(data)
            u(1) = BitConverter.ToUInt16(b, 0)
            u(0) = BitConverter.ToUInt16(b, 2)
            Return u

        End Function

        ''' <summary>
        ''' 一个单精度数组转换为Ushort数组
        ''' </summary>
        ''' <param name="data"></param>
        ''' <returns></returns>
        Public Function ArraySingleToUshort(data As Single()) As UShort()

            Dim lenth As Integer = data.Length
            Dim u(2 * lenth - 1) As UShort
            Dim b As Byte()

            For i As Integer = 0 To lenth - 1
                b = BitConverter.GetBytes(data(i))
                u(2 * i + 1) = BitConverter.ToUInt16(b, 0)
                u(2 * i) = BitConverter.ToUInt16(b, 2)
            Next

            Return u

        End Function
        ''' <summary>
        ''' 有符号16位整形，转Ushort
        ''' </summary>
        ''' <param name="data"></param>
        ''' <returns></returns>
        Public Function Int16ToUshort(data As Int16) As UShort

            Dim a As Byte() = BitConverter.GetBytes(data)

            Return BitConverter.ToUInt16(a, 0)

        End Function
    End Class

    ''' <summary>
    ''' VOC系统中，需要保存到报表的参数
    '''  conc_h 采集仪器得到湿基浓度,conc 计算得到的干基浓度,conc_z 计算得到的折算浓度,conc_p 计算得到的排放浓度
    ''' </summary>
    Class VOCParameterName
        Public ch4conc_h As Single
        Public ch4conc As Single
        Public ch4conc_z As Single
        Public ch4conc_p As Single

        Public nmhcconc_h As Single
        Public nmhcconc As Single
        Public nmhcconc_z As Single
        Public nmhcconc_p As Single

        Public tchconc_h As Single
        Public tchconc As Single
        Public tchconc_z As Single
        Public tchconc_p As Single

        '颗粒物    add by wx 2024.08.20 15:19
        Public particulateconc_h As Single
        Public particulateconc As Single
        Public particulateconc_z As Single
        Public particulateconc_p As Single



        Public o2conc_h As Single
        Public o2conc As Single

        '温压流监测仪
        Public tempconc As Single
        Public pressconc As Single
        Public speedconc As Single
        Public flowconc_g As Single
        Public flowconc_b As Single

        '湿度
        Public humiconc As Single
        Public fuheconc As Single

        Public note As String
    End Class

    ''' <summary>
    ''' VOC系统中的关键参数
    ''' </summary>
    Class VOCKeyParameter
        Public area As Single    '烟道面积 m2
        Public codffO2 As Single '空气系数
        Public speed As Single   '速度场系数
        Public pitot As Single
        Public press As Single   '大气压
        Public limitO2 As Single '折算氧限值
        Public corverFlag As Single '是否折算
    End Class
    '测量参数量程设置
    Class VOCRangeSetting
        Public tchup As Single  '总烃量程上限值
        Public tchlow As Single  '总烃量程下限值
        Public tchalarmup As Single  '总烃报警上限值
        Public tchalarmlow As Single '总烃报警下限值

        Public ch4up As Single  '甲烷量程上限值
        Public ch4low As Single '甲烷量程下限值
        Public ch4alarmup As Single  '甲烷报警上限值
        Public ch4alarmlow As Single '甲烷报警下限值

        Public nmhcup As Single  '非甲烷总烃量程上限值
        Public nmhclow As Single '非甲烷总烃量程下限值
        Public nmhcalarmup As Single '非甲烷总烃报警上限值
        Public nmhcalarmlow As Single  '非甲烷总烃报警下限值

        Public o2up As Single   '氧气量程上限
        Public o2low As Single   '氧气量程下限值
        Public o2alarmup As Single '氧气报警上限值
        Public o2alarmlow As Single '氧气报警下限值

        Public temprangeup As Single '烟温 量程上限
        Public temprangelow As Single '烟温 量程下限
        Public pressrangeup As Single '静压 量程上限
        Public pressrangelow As Single '静压 量程下限
        '此量程的设置是为了计算流速（如果流速需要上位机计算）
        Public speedrangeup As Single '动压[限值] 量程上限
        Public speedrangelow As Single '动压[限值] 量程下限
        Public humirangeup As Single  '湿度量程上限值
        Public humirangelow As Single '湿度量程下限值

        Public particulaterangeup As Single  '颗粒物量程上限值
        Public particulaterangelow As Single '颗粒物量程下限值
        Public particulatealarmup As Single  '颗粒物量程报警上限值
        Public particulatealarmlow As Single '颗粒物量程报警下限值



    End Class
    '测量参数报警设置
    Class AlarmSetting
        Public ch4alarmlow As Single
        Public ch4alarmup As Single
        Public ch4rangelow As Single
        Public ch4rangeup As Single

        Public tchalarmlow As Single
        Public tchalarmup As Single
        Public tchrangelow As Single
        Public tchrangeup As Single

        Public nmhcalarmlow As Single
        Public nmhcalarmup As Single
        Public nmhcrangelow As Single
        Public nmhcrangeup As Single

        '颗粒物
        Public particulatealarmlow As Single
        Public particulatealarmup As Single
        Public particulaterangelow As Single
        Public particulaterangeup As Single


    End Class

    '反吹设置
    Class PurgeGasSetting
        Public purgeFlag As Boolean     '是否反吹,   1 反吹,0 禁用反吹
        Public gasPurgeTime As Integer  '烟气反吹时间间隔 h
        Public gasPurgeHold As Integer  '烟气反吹 工作持续时间s
        Public pitotTime As Integer     '皮托管反吹时间间隔 h
        Public pitotHold As Integer     '皮托管反吹 工作持续时间s
    End Class

    '校准设置
    Class CalibrationSetting
        Public zeroTime As Integer   '校零 时长(秒)
        Public zeroConc As Single    '校零 浓度(ms/m3)
        Public rangeTime As Integer  '校标 时长(秒)
        Public rangeConc As Single   '校标 浓度(ms/m3)
        Public calibrationTime As Integer '执行校零后定时复位（秒）
        Public rangeResetTime As Integer  '执行校标后定时复位（秒）
        Public platformResetTime As Integer '平台阀打开后定时复位（秒）
    End Class

    Class UserInfomation
        Public userName As String
        Public userType As Integer
        Public userPW As String
    End Class
    ''' <summary>
    ''' CRC_16校验
    ''' </summary>
    ''' <param name="Coun">数组名</param>
    ''' <param name="lowIndex">数组的下标</param>
    ''' <param name="upIndex">数组上标</param>
    ''' <returns></returns>
    Public Function ModbusCRC_16(Coun() As Byte, lowIndex As Integer, upIndex As Integer) As Long
        Dim intBit, intTemp As Integer
        Dim lonCRC As Long
        Dim intCnt As Long

        lonCRC = &HFFFF&
        For intCnt = lowIndex To upIndex
            lonCRC = lonCRC Xor Coun(intCnt)
            For intBit = 0 To 7
                intTemp = lonCRC Mod 2
                lonCRC = lonCRC \ 2
                If intTemp = 1 Then
                    lonCRC = lonCRC Xor &HA001&
                End If
            Next intBit
        Next intCnt
        Return lonCRC
    End Function
End Module
