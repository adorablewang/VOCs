Imports Modbus.Device
Imports System.Threading
Imports Ini.Net
Imports System.Text
Imports System.ComponentModel
Imports System.IO.Ports
Imports Modbus.Utility
Imports System.IO
Imports System.Net.Sockets
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement



Public Class FormMain


    Private master As ModbusSerialMaster

    '发送实时数据时间间隔
    Private SendRTDataTime As Integer = 30

    '程序是否是第一次启动
    Public IniSoftStart As Boolean = True
    Public oldSaveTime As Integer
    '创建委托
    Private Delegate Sub displayConc()
    '声明委托
    Private myDisplay As displayConc
    '用来刷新时钟
    Private WithEvents Timer1 As System.Windows.Forms.Timer
    '用了刷新瞬时数据
    Private WithEvents Timer2 As System.Windows.Forms.Timer
    '发送实时数据，临时使用
    Private WithEvents Timer4 As System.Windows.Forms.Timer
    '系统是否开始校准,当为1时，系统校零，为2时，系统校标，为0时，正常采样
    Private CalOrZeroFlag As Integer

    Private WithEvents serialtest As New SerialPort
    Private tchtest As Single
    Private ch4test As Single
    Private nmhctest As Single

    '全部的实时曲线显示的点
    Private Const pointCount As Int32 = 200
    Private listx As New List(Of String)
    Private list1 As New List(Of Single)
    Private list2 As New List(Of Single)
    Private list3 As New List(Of Single)
    Private list4 As New List(Of Single)
    Private list5 As New List(Of Single)
    Private list6 As New List(Of Single)
    Private list7 As New List(Of Single)

    '经过计算后的分钟数据，数据来源是每5秒的数据累加后取平均
    Private saveMinuteAverage As New VOCParameterName
    Private calOrzerobuff As String '校0，校标
    Private calPlatformValve As String  ' 平台校准阀  关 N, 开 C

    '记录用户首次登录时，登录对话框的返回值
    Public loginVerification As Boolean = True

#Region "申明用户控件，显示主测参数"
    Friend WithEvents Usercontrol_ch4 As New UserControl_One(3, "甲烷(标干)")
    Friend WithEvents Usercontrol_Tch As New UserControl_One(3, "总烃(标干)")
    Friend WithEvents Usercontrol_Nmhc As New UserControl_One(3, "非甲烷总烃(标干)")
    Friend WithEvents Usercontrol_O2 As New UserControl_One(1, "氧气(标干)")

    Friend WithEvents Usercontrol_Particulate As New UserControl_One(3, "颗粒物(标干)") ' add by wx at 2024.08.19 10:40


    Friend WithEvents Usercontrol_TEMP As New UserControl_One(1, "温度")
    Friend WithEvents Usercontrol_PRESS As New UserControl_One(1, "压力")
    Friend WithEvents Usercontrol_SPEED As New UserControl_One(3, "流速")
    Friend WithEvents Usercontrol_HUMI As New UserControl_One(1, "湿度")

    Friend WithEvents Usercontrol_ch4_h As New UserControl_One(1, "甲烷(湿基)")
    Friend WithEvents Usercontrol_tch_h As New UserControl_One(1, "总烃(湿基)")
    Friend WithEvents Usercontrol_nmhc_h As New UserControl_One(1, "非甲烷总烃(湿基)")
    Friend WithEvents Usercontrol_o2_h As New UserControl_One(1, "氧气(湿基)")

    Friend WithEvents Usercontrol_Particulate_h As New UserControl_One(1, "颗粒物(湿基)") ' add by wx at 2024.08.19 10:40


    Friend WithEvents Temperature1 As New ParameterDisplay
    Friend WithEvents Temperature2 As New ParameterDisplay
    Friend WithEvents Temperature3 As New ParameterDisplay
#End Region

#Region "分钟数据导入小时数据，小时数据导入日数据需要用到的字符串"
    Private ReadOnly sqlAvgString As String = "round(avg(ch4conc_h),2) as ch4conc_h,round(avg(ch4conc),2) as ch4conc,round(avg(ch4conc_z),2) as ch4conc_z,round(avg(ch4conc_p),2) as ch4conc_p," &
            "round(avg(nmhcconc_h),2) as nmhcconc_h,round(avg(nmhcconc),2) as nmhcconc,round(avg(nmhcconc_z),2) as nmhcconc_z,round(avg(nmhcconc_p),2) as nmhcconc_p," &
            "round(avg(tchconc_h),2) as tchconc_h,round(avg(tchconc),2) as tchconc,round(avg(tchconc_z),2) as tchconc_z,round(avg(tchconc_p),2) as tchconc_p," &
            "round(avg(particulateconc_h),2) as particulateconc_h,round(avg(particulateconc),2) as particulateconc,round(avg(particulateconc_z),2) as particulateconc_z,round(avg(particulateconc_p),2) as particulateconc_p," &
            "round(avg(O2Conc_h),2) as O2Conc_h,round(avg(O2Conc),2) as O2Conc," &
            "round(avg(TEMPConc),2) as TEMPConc," &
            "round(avg(PressConc),2) as PressConc," &
            "round(avg(speedconc),2) as speedconc," &
            "round(avg(flowconc_g),2) as SpeedConc," &
            "round(avg(flowconc_b),2) as FlowConc," &
            "round(avg(humiconc),2) as humiconc," &
            "0.0 as huheconc," &
            "0 as remark "

    Private ReadOnly sqlDayString As String = "round(avg(ch4conc_h),2) as ch4conc_h,round(avg(ch4conc),2) as ch4conc,round(avg(ch4conc_z),2) as ch4conc_z,round(sum(ch4conc_p),2) as ch4conc_p," &
            "round(avg(nmhcconc_h),2) as nmhcconc_h,round(avg(nmhcconc),2) as nmhcconc,round(avg(nmhcconc_z),2) as nmhcconc_z,round(sum(nmhcconc_p),2) as nmhcconc_p," &
            "round(avg(tchconc_h),2) as tchconc_h,round(avg(tchconc),2) as tchconc,round(avg(tchconc_z),2) as tchconc_z,round(sum(tchconc_p),2) as tchconc_p," &
            "round(avg(particulateconc_h),2) as particulateconc_h,round(avg(particulateconc),2) as particulateconc,round(avg(particulateconc_z),2) as particulateconc_z,round(avg(particulateconc_p),2) as particulateconc_p," &
            "round(avg(O2Conc_h),2) as O2Conc_h,round(avg(O2Conc),2) as O2Conc," &
            "round(avg(TEMPConc),2) as TEMPConc," &
            "round(avg(PressConc),2) as PressConc," &
            "round(avg(speedconc),2) as speedconc," &
            "round(sum(flowconc_g),2) as SpeedConc," &
            "round(sum(flowconc_b),2) as FlowConc," &
            "round(avg(humiconc),2) as humiconc," &
            "0.0 as huheconc," &
            "0 as remark "
#End Region

#Region "加载主测参数控件，并初始化"
    Public Sub InitializeMainPara()

        Dim strbuf As String

        Usercontrol_ch4.Tag = 1
        Usercontrol_Tch.Tag = 2
        Usercontrol_Nmhc.Tag = 3
        Usercontrol_Particulate.Tag = 4 ' 固体颗粒物(标干) add by wx at 2024.08.19 10:51
        Usercontrol_O2.Tag = 5
        Usercontrol_TEMP.Tag = 6
        Usercontrol_PRESS.Tag = 7
        Usercontrol_SPEED.Tag = 8
        Usercontrol_HUMI.Tag = 9

        Usercontrol_ch4_h.Tag = 10
        Usercontrol_tch_h.Tag = 11
        Usercontrol_nmhc_h.Tag = 12
        Usercontrol_Particulate_h.Tag = 13 ' 固体颗粒物(湿基) add by wx at 2024.08.19 10:51
        Usercontrol_o2_h.Tag = 14

        FlowLayoutPanel1.Controls.Add(Usercontrol_ch4)
        FlowLayoutPanel1.Controls.Add(Usercontrol_Tch)
        FlowLayoutPanel1.Controls.Add(Usercontrol_Nmhc)
        FlowLayoutPanel1.Controls.Add(Usercontrol_Particulate) ' 固体颗粒物(标干) add by wx at 2024.08.19 10:51
        FlowLayoutPanel1.Controls.Add(Usercontrol_O2)
        FlowLayoutPanel1.Controls.Add(Usercontrol_TEMP)
        FlowLayoutPanel1.Controls.Add(Usercontrol_PRESS)
        FlowLayoutPanel1.Controls.Add(Usercontrol_SPEED)
        FlowLayoutPanel1.Controls.Add(Usercontrol_HUMI)

        FlowLayoutPanel1.Controls.Add(Usercontrol_ch4_h)
        FlowLayoutPanel1.Controls.Add(Usercontrol_tch_h)
        FlowLayoutPanel1.Controls.Add(Usercontrol_nmhc_h)
        FlowLayoutPanel1.Controls.Add(Usercontrol_Particulate_h) ' 固体颗粒物(湿基) add by wx at 2024.08.19 10:51

        FlowLayoutPanel1.Controls.Add(Usercontrol_o2_h)

        Dim newstring As String = UnicodeSuperscriptConversion(3)
        Usercontrol_ch4.SetTypeUnit("浓度,折算,排量", "mg/m" & newstring & ",mg/m" & newstring & ",kg/h")
        Usercontrol_Tch.SetTypeUnit("浓度,折算,排量", "mg/m" & newstring & ",mg/m" & newstring & ",kg/h")
        Usercontrol_Nmhc.SetTypeUnit("浓度,折算,排量", "mg/m" & newstring & ",mg/m" & newstring & ",kg/h")
        'Usercontrol_Particulate.SetTypeUnit("浓度,折算,排量", "mg/m" & newstring & ",mg/m" & newstring & ",kg/h") ' 固体颗粒物(标干) add by wx at 2024.08.19 10:51

        Usercontrol_Particulate.SetTypeUnit("浓度,折算,排量", "mg/m" & newstring & ",mg/m" & newstring & ",kg/h")

        Usercontrol_SPEED.SetTypeUnit("流速,工况,标况", "m/s,m" & newstring & "/h,m" & newstring & "/h")
        Usercontrol_O2.SetTypeUnit("浓度", "%")
        Usercontrol_TEMP.SetTypeUnit("温度", "℃")
        Usercontrol_PRESS.SetTypeUnit("压力", "Pa")
        Usercontrol_HUMI.SetTypeUnit("湿度", "%")

        Usercontrol_ch4_h.SetTypeUnit("浓度", "mg/m" & newstring)
        Usercontrol_tch_h.SetTypeUnit("浓度", "mg/m" & newstring)
        Usercontrol_nmhc_h.SetTypeUnit("浓度", "mg/m" & newstring)
        Usercontrol_Particulate_h.SetTypeUnit("浓度", "mg/m" & newstring) ' 固体颗粒物(湿基) add by wx at 2024.08.19 10:51



        Usercontrol_o2_h.SetTypeUnit("浓度", "%")


        Temperature1.AnalyzerColor = Color.DodgerBlue
        Temperature1.AnalyzerSetValue = ReadConfigIni("温度设置", "温度1", "120")
        Temperature1.AnalyzerName = "采样管1温度"
        Temperature2.AnalyzerColor = Color.DarkGreen
        Temperature2.AnalyzerSetValue = ReadConfigIni("温度设置", "温度2", "100")
        Temperature2.AnalyzerName = "加热箱温度"
        Temperature3.AnalyzerColor = Color.LightGray
        Temperature3.AnalyzerSetValue = ReadConfigIni("温度设置", "温度3", "100")
        Temperature3.AnalyzerName = "采样管2温度"
        FlowLayoutPanel2.Controls.Add(Temperature1)
        FlowLayoutPanel2.Controls.Add(Temperature2)
        FlowLayoutPanel2.Controls.Add(Temperature3)

        calPlatformValve = "N" '默认时没有打开平台校准阀门，此时显示是采样状态
        strbuf = ReadConfigIni("界面显示", "显示参数", "1,1,1,1,1,1,1,1,1,1,1,1,1,1")
        displaySelect(strbuf)
    End Sub

#End Region

    '''' <summary>
    '''' 应用程序启动时添加登录验证对话框
    '''' </summary>
    'Public Sub New()
    '    '在加载FormMain之前弹出登录框
    '    Dim loginFrom As New FormLogin
    '    '以模态形式形式显示登录窗口
    '    loginFrom.ShowDialog()

    '    If loginFrom.IsLoggedIn Then
    '        '记录首次登录状态，用于判断是否加载FormMain，还是退出应用程序
    '        loginVerification = True

    '        ' 此调用是设计器所必需的。
    '        InitializeComponent()
    '    Else
    '        '登录失败，不显示主窗口
    '        '释放登录窗口资源
    '        loginFrom.Dispose()
    '    End If
    '    ' 在 InitializeComponent() 调用之后添加任何初始化。
    'End Sub

    Public Sub New()
        ' 此调用是设计器所必需的。
        InitializeComponent()

    End Sub

    ''' <summary>
    ''' 初始化设备的通讯端口
    ''' </summary>
    Private Sub InitializeSerialModbus()
        Dim strbuf As String

        strbuf = ReadConfigIni("通讯设置", "甲烷", "COM1,9600,N,8,1")
        analyzerdata.SerialConfig = strbuf

        strbuf = ReadConfigIni("通讯设置", "PLC", "COM2,9600,N,8,1")
        plcdata.SerialConfig = strbuf

        strbuf = ReadConfigIni("通讯设置", "氧气", "COM3,9600,N,8,1")
        o2data.SerialConfig = strbuf

        strbuf = ReadConfigIni("通讯设置", "颗粒物", "COM4,9600,N,8,1")
        particulatedata.SerialConfig = strbuf

        strbuf = ReadConfigIni("通讯设置", "数采仪", "COM5,9600,N,8,1")
        InitializeDataSamplingAnalyzer(strbuf)
    End Sub

    Private Sub InitializeDataSamplingAnalyzer(strbuf As String)
        Dim tokens As String()

        tokens = strbuf.Split(",")
        datasamplinganalyzer.PortName = tokens(0)
        datasamplinganalyzer.BaudRate = tokens(1)

        If tokens(2).ToUpper = "N" Then
            datasamplinganalyzer.Parity = Parity.None
        ElseIf tokens(2).ToUpper = "O" Then
            datasamplinganalyzer.Parity = Parity.Odd
        ElseIf tokens(2).ToUpper = "E" Then
            datasamplinganalyzer.Parity = Parity.Even
        End If

        datasamplinganalyzer.DataBits = tokens(3)

        If tokens(4).ToUpper = "1" Then
            datasamplinganalyzer.StopBits = StopBits.One
        ElseIf tokens(4).ToUpper = "1.5" Then
            datasamplinganalyzer.StopBits = StopBits.OnePointFive
        ElseIf tokens(4).ToUpper = "2" Then
            datasamplinganalyzer.StopBits = StopBits.Two
        End If
        Try
            datasamplinganalyzer.Open()
        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' 初始化系统的关键参数，包括烟道面积、空气系数、皮托管系数、速度场系数
    ''' </summary>
    Private Sub InitializeKeyParameter()
        Dim strbuf As String
        Dim tokens As String()

        strbuf = ReadConfigIni("关键参数", "烟道面积", "10")
        keyParameter.area = strbuf

        strbuf = ReadConfigIni("关键参数", "大气压", "101325")
        keyParameter.press = strbuf

        strbuf = ReadConfigIni("关键参数", "空气系数", "1.4")
        keyParameter.codffO2 = strbuf

        strbuf = ReadConfigIni("关键参数", "速度场系数", "1")
        keyParameter.speed = strbuf

        strbuf = ReadConfigIni("关键参数", "皮托管系数", "0.84")
        keyParameter.pitot = strbuf

        strbuf = ReadConfigIni("关键参数", "折算氧限值", "18")
        keyParameter.limitO2 = strbuf

        strbuf = ReadConfigIni("关键参数", "是否折算", "1")
        keyParameter.corverFlag = strbuf

        strbuf = ReadConfigIni("反吹设置", "反吹时间设置", "1,2,300,2,60")
        tokens = strbuf.Split(",")
        purgeSetting.purgeFlag = tokens(0)
        purgeSetting.gasPurgeTime = tokens(1)
        purgeSetting.gasPurgeHold = tokens(2)
        purgeSetting.pitotTime = tokens(3)
        purgeSetting.pitotHold = tokens(4)

    End Sub

    Private Sub InitializePurgeAndCalibration()
        Dim strbuf As String
        Dim tokens As String()

        strbuf = ReadConfigIni("反吹设置", "反吹时间设置", "1,2,300,2,60")
        tokens = strbuf.Split(",")
        purgeSetting.purgeFlag = tokens(0)  '是否反吹,   1 反吹,0 禁用反吹
        purgeSetting.gasPurgeTime = tokens(1) '烟气反吹时间间隔 h
        purgeSetting.gasPurgeHold = tokens(2) '烟气反吹 工作持续时间s
        purgeSetting.pitotTime = tokens(3) '皮托管反吹时间间隔 h
        purgeSetting.pitotHold = tokens(4) '皮托管反吹 工作持续时间s

        '校准复位定时器默认时间2700秒（45分钟）
        strbuf = ReadConfigIni("校准设置", "校准参数", "360,0,360,20,2700,2700,2700")
        tokens = strbuf.Split(",")
        calibrationPara.zeroTime = tokens(0)  '校零 时长(秒)
        calibrationPara.zeroConc = tokens(1)  '校零 浓度(ms/m3)
        calibrationPara.rangeTime = tokens(2) '校标 时长(秒)
        calibrationPara.rangeConc = tokens(3) '校标 浓度(ms/m3)
        '校零或校标定时器的时长
        calibrationPara.calibrationTime = tokens(4)
        calibrationPara.rangeResetTime = tokens(5)
        calibrationPara.platformResetTime = tokens(6)

    End Sub


    ''' <summary>
    ''' 初始化测量参数的上下限值
    ''' </summary>
    Private Sub InitializeParameterRange()
        Dim strbuf As String
        Dim tokens As String()

        strbuf = ReadConfigIni("量程设置", "温度限值", "300,0,300,0")
        tokens = strbuf.Split(",")
        rangeParameter.temprangeup = tokens(0)
        rangeParameter.tchalarmlow = tokens(1)

        strbuf = ReadConfigIni("量程设置", "压力限值", "10000,-10000,10000,-10000")
        tokens = strbuf.Split(",")
        rangeParameter.pressrangeup = tokens(0)
        rangeParameter.pressrangelow = tokens(1)

        strbuf = ReadConfigIni("量程设置", "动压限值", "500,0,500,0")
        tokens = strbuf.Split(",")
        rangeParameter.speedrangeup = tokens(0)
        rangeParameter.speedrangelow = tokens(1)

        strbuf = ReadConfigIni("量程设置", "湿度限值", "40,0,40,0")
        tokens = strbuf.Split(",")
        rangeParameter.humirangeup = tokens(0)
        rangeParameter.humirangelow = tokens(1)

        strbuf = ReadConfigIni("量程设置", "总烃限值", "200,0,200,0")
        tokens = strbuf.Split(",")
        rangeParameter.tchup = tokens(0)
        rangeParameter.tchlow = tokens(1)
        rangeParameter.tchalarmup = tokens(2)
        rangeParameter.tchalarmlow = tokens(3)

        strbuf = ReadConfigIni("量程设置", "甲烷限值", "200,0,200,0")
        tokens = strbuf.Split(",")
        rangeParameter.ch4up = tokens(0)
        rangeParameter.ch4low = tokens(1)
        rangeParameter.ch4alarmup = tokens(2)
        rangeParameter.ch4alarmlow = tokens(3)

        strbuf = ReadConfigIni("量程设置", "非甲烷限值", "200,0,200,0")
        tokens = strbuf.Split(",")
        rangeParameter.nmhcup = tokens(0)
        rangeParameter.nmhclow = tokens(1)
        rangeParameter.nmhcalarmup = tokens(2)
        rangeParameter.nmhcalarmlow = tokens(3)

        strbuf = ReadConfigIni("量程设置", "氧气限值", "25,0,25,0")
        tokens = strbuf.Split(",")
        rangeParameter.o2up = tokens(0)
        rangeParameter.o2low = tokens(1)
        rangeParameter.o2alarmup = tokens(2)
        rangeParameter.o2alarmlow = tokens(3)

        strbuf = ReadConfigIni("量程设置", "颗粒物限值", "90,0,90,0")
        tokens = strbuf.Split(",")
        rangeParameter.particulaterangeup = tokens(0)
        rangeParameter.particulaterangelow = tokens(1)
        rangeParameter.particulatealarmup = tokens(2)
        rangeParameter.particulatealarmlow = tokens(3)


    End Sub

    Private Sub InitializeAlarm()
        Dim strbuf As String
        Dim tokens As String()

        strbuf = ReadConfigIni("报警设置", "甲烷", "0,100,0,200")
        tokens = strbuf.Split(",")
        alarmParameter.ch4alarmlow = tokens(0)
        alarmParameter.ch4alarmup = tokens(1)
        alarmParameter.ch4rangelow = tokens(2)
        alarmParameter.ch4rangeup = tokens(3)

        strbuf = ReadConfigIni("报警设置", "总烃", "0,100,0,200")
        tokens = strbuf.Split(",")
        alarmParameter.tchalarmlow = tokens(0)
        alarmParameter.tchalarmup = tokens(1)
        alarmParameter.tchrangelow = tokens(2)
        alarmParameter.tchrangeup = tokens(3)

        strbuf = ReadConfigIni("报警设置", "非甲烷总烃", "0,100,0,200")
        tokens = strbuf.Split(",")
        alarmParameter.nmhcalarmlow = tokens(0)
        alarmParameter.nmhcalarmup = tokens(1)
        alarmParameter.nmhcrangelow = tokens(2)
        alarmParameter.nmhcrangeup = tokens(3)

        strbuf = ReadConfigIni("报警设置", "颗粒物", "0,50,0,90")
        tokens = strbuf.Split(",")
        alarmParameter.particulatealarmlow = tokens(0)
        alarmParameter.particulatealarmup = tokens(1)

    End Sub

    ''' <summary>
    ''' 初始化上传到服务器通讯协议，或者为实验室检测，或者为现场检验，为0时候，实验室认证，为1时候，现场认证
    ''' </summary>
    Private Sub InitializeProtocol()
        Dim strbuf As String

        strbuf = ReadConfigIni("通讯协议", "协议选择", "0")
        SelectProtocol = strbuf

    End Sub

    Private Sub Temperature1_ChangeColor() Handles Temperature1.ChangeColor
        Dim colordialog As New ColorDialog
        If colordialog.ShowDialog = DialogResult.OK Then Temperature1.AnalyzerColor = colordialog.Color
    End Sub

    Private Sub Temprature1_ChangeSetValue() Handles Temperature1.ChangeSetValue
        Dim input As String = InputBox("请输入设定温度", "设定温度")
        Dim InitializeName As New IniFile(IniPathName)

        If input = "" Then

        Else
            Temperature1.AnalyzerSetValue = input

            Dim buff As UShort()
            ReDim buff(0)

            buff(0) = CUShort(input)

            plcdata.writeCom = True
            plcdata.writeAnaly = 1
            plcdata.writeAdd = 70
            plcdata.writeData = buff

            InitializeName.WriteString("温度设置", "温度1", input)
        End If
    End Sub

    Private Sub Temperature2_ChangeColor() Handles Temperature2.ChangeColor
        Dim colordialog As New ColorDialog
        If colordialog.ShowDialog = DialogResult.OK Then Temperature2.AnalyzerColor = colordialog.Color
    End Sub

    Private Sub Temprature2_ChangeSetValue() Handles Temperature2.ChangeSetValue
        Dim input As String = InputBox("请输入设定温度", "设定温度")
        Dim InitializeName As New IniFile(IniPathName)

        If input = "" Then

        Else
            Temperature2.AnalyzerSetValue = input

            Dim buff As UShort()
            ReDim buff(0)

            buff(0) = CUShort(input)

            plcdata.writeCom = True
            plcdata.writeAnaly = 1
            plcdata.writeAdd = 71
            plcdata.writeData = buff

            InitializeName.WriteString("温度设置", "温度2", input)
        End If
    End Sub

    Private Sub Temperature3_ChangeColor() Handles Temperature3.ChangeColor
        Dim colordialog As New ColorDialog
        If colordialog.ShowDialog = DialogResult.OK Then Temperature3.AnalyzerColor = colordialog.Color
    End Sub

    Private Sub Temprature3_ChangeSetValue() Handles Temperature3.ChangeSetValue
        Dim input As String = InputBox("请输入设定温度", "设定温度")
        Dim InitializeName As New IniFile(IniPathName)

        If input = "" Then

        Else
            Temperature3.AnalyzerSetValue = input

            Dim buff As UShort()
            ReDim buff(0)

            buff(0) = CUShort(input)

            plcdata.writeCom = True
            plcdata.writeAnaly = 1
            plcdata.writeAdd = 72
            plcdata.writeData = buff

            InitializeName.WriteString("温度设置", "温度3", input)
        End If
    End Sub

    Private Sub FormMain_Load(sender As Object, e As EventArgs) Handles Me.Load

        '登录对话框没有校验成功，或者用户点击“取消”
        If Not loginVerification Then
            Environment.Exit(0)
        End If

        FlowLayoutPanel1.Dock = DockStyle.Fill
        FlowLayoutPanel2.Dock = DockStyle.Fill
        Panel1.Dock = DockStyle.Fill
        SplitContainer2.Dock = DockStyle.Fill
        Chart1.Dock = DockStyle.Fill
        '实例化委托
        myDisplay = AddressOf DisplayAllParameter

        '程序是否第一次启动，如果是第一次启动，给保存数据的加一个标志
        If IniSoftStart Then oldSaveTime = Now.Minute
        '初始化串口
        InitializeSerialModbus()
        '初始化关键参数
        InitializeKeyParameter()
        '初始化量程限值
        InitializeParameterRange()
        '初始化报警参数
        InitializeAlarm()
        '初始化与服务器通讯协议
        InitializeProtocol()
        '初始化校准参数和反吹参数
        InitializePurgeAndCalibration()
        '暂时不使用TCP和PLC通讯
        'plctcp = New PLC_TCP("127.0.0.1,5000")
        '启动 分析仪,plc ，氧气采集仪,颗粒物采集仪开始工作，包括串口通讯建立和 启动后台子线程采集运行
        analyzerdata.InitializePara()
        plcdata.InitializePara()
        o2data.InitializePara()
        particulatedata.InitializePara()

        ' 加一个 颗粒物

        '委托，用来显示主界面
        maindisplay = AddressOf displaySelect
        InitializeMainPara()

        '定时器1s 更新一次窗体时间
        Timer1 = New Windows.Forms.Timer
        Timer1.Interval = 1000
        Timer1.Start()

        '定时器 五秒获取保存一次测量参数，用于计算分钟平均值
        Timer2 = New Windows.Forms.Timer
        Timer2.Interval = 5000
        Timer2.Start()

        '60s发送一次 实时数据(当前时刻各个采集指标的瞬时值)
        '5s
        Timer4 = New Windows.Forms.Timer
        Timer4.Interval = 1000 * 5
        Timer4.Start()


        '临时添加，和公司设备调试使用
        'serialtest.PortName = "COM6"
        'serialtest.BaudRate = 115200
        'serialtest.Parity = Parity.None
        'serialtest.DataBits = 8
        'serialtest.StopBits = StopBits.One
        'serialtest.Open()

        InitializeMN()
        WriteLoginInfo()
        InsertAdmin()
    End Sub

    Private Sub WriteLogAlarm()
        Dim sqlHelper As SQLiteHelper
        Dim list1 As New ArrayList

        Dim note As String = "N"

        sqlHelper = New SQLiteHelper(ConfigFilePath)
        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        '软件启动时，记录启动时间到数据库中
        list1.Add("insert into 报警记录(samplingtime ,info ,infotype) values('" & Now.ToString("yyyy-MM-dd HH:mm:ss") & "','折算浓度报警',1);")
        sqlHelper.IUD(list1)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If
    End Sub

    ''' <summary>
    ''' 程序启动时，写登录日志
    ''' </summary>
    Private Sub WriteLoginInfo()
        Dim sqlHelper As SQLiteHelper
        Dim list1 As New ArrayList
        Dim str As String

        Dim note As String = "N"
        Dim userInfo As String = ""

        sqlHelper = New SQLiteHelper(ConfigFilePath)
        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        '记录用户类型
        If UserRole = EnumDefine.UserRoles.Administrator Then
            userInfo = "管理员"
        ElseIf UserRole = EnumDefine.UserRoles.Inspector Then
            userInfo = "巡视员"
        End If

        Dim tokens As String()

        '判断时候有意外断电或者强制退出文件，
        If My.Computer.FileSystem.FileExists(Application.StartupPath & "\log\power\power.txt") Then
            Dim sr As StreamReader = New StreamReader(Application.StartupPath & "\log\power\power.txt")

            str = sr.ReadLine.Trim
            '当字符串中有连续两个空格，删除一个
            'Do While InStr(str, "  ") <> 0
            '    str = str.Replace("  ", " ")
            'Loop
            '使用正则表达式，去除字符串中的两个以上的空格
            str = RegularExpressions.Regex.Replace(str, " {2,}", " ")

            tokens = str.Split(" ")
            sr.Close()

            If tokens.Length > 2 Then
                str = tokens(1) & " " & tokens(2)
                list1.Add("insert into 事件记录(samplingtime,username,info,infotype) values('" & str & "','" & userInfo & "','系统断电或者软件被强制退出',1);")
            End If
        End If

        '软件启动时，记录启动时间到数据库中
        list1.Add("insert into 事件记录(samplingtime,username,info,infotype) values('" & Now.ToString("yyyy-MM-dd HH:mm:ss") & "','" & userInfo & "','系统软件启动',1);")
        sqlHelper.IUD(list1)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If
    End Sub

    Private Sub WriteCalLog(strinfo As String)
        Dim sqlHelper As SQLiteHelper
        Dim list1 As New ArrayList
        Dim str As String = Now.ToString("yyyy-MM-dd HH:mm:ss")

        Dim note As String = "N"
        Dim userInfo As String = ""

        sqlHelper = New SQLiteHelper(ConfigFilePath)
        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        '记录用户类型
        If UserRole = EnumDefine.UserRoles.Administrator Then
            userInfo = "管理员"
        ElseIf UserRole = EnumDefine.UserRoles.Inspector Then
            userInfo = "巡视员"
        End If

        '记录校准记录
        list1.Add("insert into 事件记录(samplingtime,username,info,infotype) values('" & str & "','" & userInfo & "','" & strinfo & "',2);")
        sqlHelper.IUD(list1)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If
    End Sub
    Private Sub WriteLogOutInfo()
        Dim sqlHelper As SQLiteHelper
        Dim list1 As New ArrayList

        Dim note As String = "N"
        Dim userInfo As String = ""

        sqlHelper = New SQLiteHelper(ConfigFilePath)
        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        '记录用户类型
        If UserRole = EnumDefine.UserRoles.Administrator Then
            userInfo = "管理员"
        ElseIf UserRole = EnumDefine.UserRoles.Inspector Then
            userInfo = "巡视员"
        End If

        list1.Add("insert into 事件记录(samplingtime,username,info,infotype) values('" & Now.ToString("yyyy-MM-dd HH:mm:ss") & "','" & userInfo & "','系统软件退出',1)")
        sqlHelper.IUD(list1)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If
    End Sub

    ''' <summary>
    ''' 回调函数，控制显示主界面
    ''' </summary>
    ''' <param name="strbuf"></param>
    ''' <returns></returns>
    Private Function displaySelect(strbuf As String) As String
        Dim tokens As String()

        tokens = strbuf.Split(",")

        For Each us As UserControl_One In FlowLayoutPanel1.Controls
            us.Visible = tokens(us.Tag - 1)
        Next

        Return Convert.ToString(True)
    End Function

    Private Sub InitializeMN()
        Dim strbuf As String

        strbuf = ReadConfigIni("站点信息", "MN号", "20230212049_1")
        senddatatoserver.UserMN = strbuf

        strbuf = ReadConfigIni("站点信息", "站点名称", "湖北方圆科技")
        senddatatoserver.UserName = strbuf
    End Sub

    Private Sub serialtest_receive(sender As Object, e As SerialDataReceivedEventArgs) Handles serialtest.DataReceived

        Thread.Sleep(5)

        Dim receivedata As Byte()

        Dim count As Integer = serialtest.BytesToRead

        If count < 1 Then Exit Sub

        ReDim receivedata(count - 1)
        serialtest.Read(receivedata, 0, count)

        '判断校验，调试过程会先关闭
        'If (&HFA <> receivedata(0)) OrElse (&HFB <> receivedata(count - 1)) Then Exit Sub
        Dim m As Byte() = BitConverter.GetBytes(ModbusCRC_16(receivedata, 0, receivedata.Length - 4))
        'If (m(0) <> receivedata(count - 2)) OrElse (m(1) <> receivedata(count - 3)) Then Exit Sub

        Dim a(5) As Byte


        If receivedata(3) = 44 Then
            For i As Integer = 0 To 2
                a(2 * i) = receivedata(5 + 2 * i + 1)
                a(2 * i + 1) = receivedata(5 + 2 * i)
            Next
            tchtest = BitConverter.ToInt16(a, 0) / 100
            ch4test = BitConverter.ToInt16(a, 2) / 100
            nmhctest = BitConverter.ToInt16(a, 4) / 100
        End If
    End Sub

    Private Sub Timer4_Tick() Handles Timer4.Tick
        'If Now.Second = 20 Then
        SendRealTimeData()
        'End If

    End Sub

    Private Sub Timer1_Tick() Handles Timer1.Tick
        ToolStripStatusLabel1.Text = System.DateTime.Now.ToString("yyyy-MM-dd")
        ToolStripStatusLabel2.Text = System.DateTime.Now.ToString("T")

        '更新用户类型信息
        If UserRole = EnumDefine.UserRoles.Administrator Then
            ToolStripStatusLabel3.Text = "用户类型：管理员"
        ElseIf UserRole = EnumDefine.UserRoles.Inspector Then
            ToolStripStatusLabel3.Text = "用户类型：巡视员"
        Else
            ToolStripStatusLabel3.Text = ""
        End If
    End Sub

    Private Sub Timer2_Tick() Handles Timer2.Tick
        Static lastTriggerHour As Integer = -1
        Static iDaily As Integer
        Static iMonth As Integer
        '显示所有的浓度
        DisPlayAllConcentration()

        If oldSaveTime <> Now.Minute Then
            oldSaveTime = Now.Minute

            '计算平均值
            MinuteAverage()
            minuteAvg.Clear()

            SaveMinuteData()

            '发送分钟数据
            SendMinuteData()

        End If

        If Now.Minute = 1 AndAlso Now.Second <= 10 AndAlso lastTriggerHour <> Now.Hour Then
            '每个整点的1分钟后记录小时数据
            lastTriggerHour = Now.Hour
            InsertHourData()

            '发送小时数据
            SendHourData()
        End If

        If Now.Minute = 0 AndAlso iDaily <> Now.Day Then
            iDaily = Now.Day
            InsertDailyData()
        End If

        If Now.Hour = 0 AndAlso iMonth <> Now.Month Then
            iMonth = Now.Month
            InsertMonthData()
        End If

        WriteLogToText("power", "power", Now.ToString("yyyy-MM-dd HH:mm:ss"), False)
    End Sub

#Region "保存分钟数据、导入小时数据、导入日数据、导入月数据"

    Private Sub SaveMinuteData()
        Dim sqlHelper As SQLiteHelper
        Dim list1 As New ArrayList
        Dim str As String = Now.ToString("yyyy-MM-dd HH:mm") & ":00"

        Dim note As String = "N"

        sqlHelper = New SQLiteHelper(ConfigFilePath)
        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        list1.Add("insert into 分钟数据(samplingtime ,ch4conc_h,ch4conc,ch4conc_z,ch4conc_p
                                        ,nmhcconc_h,nmhcconc,nmhcconc_z,nmhcconc_p
                                        ,tchconc_h,tchconc,tchconc_z,tchconc_p
                                        ,particulateconc_h,particulateconc,particulateconc_z,particulateconc_p
                                        ,o2conc_h,o2conc
                                        ,tempconc ,pressconc ,speedconc ,flowconc_g,flowconc_b
                                        ,humiconc ,fuheconc,REMARK) 
                                values('" & str & "'," & saveMinuteAverage.ch4conc_h & "," & saveMinuteAverage.ch4conc & "," & saveMinuteAverage.ch4conc_z & "," & saveMinuteAverage.ch4conc_p & "
                                        ," & saveMinuteAverage.nmhcconc_h & "," & saveMinuteAverage.nmhcconc & "," & saveMinuteAverage.nmhcconc_z & "," & saveMinuteAverage.nmhcconc_p & "
                                        ," & saveMinuteAverage.tchconc_h & "," & saveMinuteAverage.tchconc & "," & saveMinuteAverage.tchconc_z & "," & saveMinuteAverage.tchconc_p & "
                                        ," & saveMinuteAverage.particulateconc_h & "," & saveMinuteAverage.particulateconc & "," & saveMinuteAverage.particulateconc_z & "," & saveMinuteAverage.particulateconc_p & "                               
                                        ," & saveMinuteAverage.o2conc_h & "," & saveMinuteAverage.o2conc & "
                                        ," & saveMinuteAverage.tempconc & "," & saveMinuteAverage.pressconc / 1000 & "," & saveMinuteAverage.speedconc & "," & saveMinuteAverage.flowconc_g & ", " & saveMinuteAverage.flowconc_b & "
                                        ," & saveMinuteAverage.humiconc & "," & keyParameter.press & ",'" & saveMinuteAverage.note & "');")

        sqlHelper.IUD(list1)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If
    End Sub
    ''' <summary>
    ''' 分钟数据插入小时数据表中
    ''' </summary>
    Private Sub InsertHourData()
        Dim startTime As String
        Dim endTime As String
        Dim sql As New StringBuilder

        Dim sqlHelper As SQLiteHelper
        Dim list1 As New ArrayList

        sqlHelper = New SQLiteHelper(ConfigFilePath)

        '首先判断是否已经有小时数据
        startTime = Now.AddHours(-1).ToString("yyyy-MM-dd HH") & ":00"
        endTime = Now.ToString("yyyy-MM-dd HH") & ":00"

        sql.Append("select * from 小时数据 where samplingtime ='" & startTime & "'")
        sqlHelper.SelectToDt(sql.ToString)
        If sqlHelper.RowsAffected > 0 Then Exit Sub

        sql.Clear()
        sql.Append("Insert into 小时数据 (")
        sql.Append("samplingtime,ch4conc_h,ch4conc,ch4conc_z,ch4conc_p,nmhcconc_h,nmhcconc,nmhcconc_z,nmhcconc_p,tchconc_h,tchconc,tchconc_z,tchconc_p,")
        sql.Append("particulateconc_h,particulateconc,particulateconc_z,particulateconc_p,")
        sql.Append("o2conc_h,o2conc,tempconc,pressconc,speedconc,flowconc_g,flowconc_b,humiconc,fuheconc,REMARK )")
        sql.Append(" Select strftime('%Y-%m-%d %H:00',samplingtime),")
        sql.Append(sqlAvgString)
        sql.Append(" from 分钟数据")
        sql.Append(" where samplingtime >= '" & startTime & "' And samplingtime < '" & endTime & "'")
        sql.Append(" group by strftime('%Y%m%d%H',samplingtime)")

        list1.Add(sql.ToString())
        sqlHelper.IUD(list1)

    End Sub

    ''' <summary>
    ''' 小时数据导入日数据库中
    ''' </summary>
    Private Sub InsertDailyData()
        Dim startTime As String
        Dim endTime As String
        Dim sql As New StringBuilder
        '首先判断是否已经有小时数据
        Dim sqlHelper As SQLiteHelper
        Dim list1 As New ArrayList

        sqlHelper = New SQLiteHelper(ConfigFilePath)

        '首先判断是否已经有日数据
        startTime = Now.AddDays(-1).ToString("yyyy-MM-dd")
        endTime = Now.ToString("yyyy-MM-dd")

        sql.Append("select * from 日数据 where samplingtime ='" & startTime & "'")
        sqlHelper.SelectToDt(sql.ToString)
        If sqlHelper.RowsAffected > 0 Then Exit Sub

        startTime = Now.AddDays(-1).ToString("yyyy-MM-dd") & " " & "00:00"

        sql.Clear()
        sql.Append("Insert into 日数据 (")
        sql.Append("samplingtime,ch4conc_h,ch4conc,ch4conc_z,ch4conc_p,nmhcconc_h,nmhcconc,nmhcconc_z,nmhcconc_p,tchconc_h,tchconc,tchconc_z,tchconc_p,")
        sql.Append("particulateconc_h,particulateconc,particulateconc_z,particulateconc_p,")
        sql.Append("o2conc_h,o2conc,tempconc,pressconc,speedconc,flowconc_g,flowconc_b,humiconc,fuheconc,REMARK )")
        sql.Append(" Select strftime('%Y-%m-%d',samplingtime),")
        sql.Append(sqlDayString)
        sql.Append(" from 小时数据")
        sql.Append(" where samplingtime >= '" & startTime & "' And samplingtime < '" & endTime & "'")
        sql.Append(" group by strftime('%Y.%m.%d',samplingtime)")

        list1.Add(sql.ToString())
        sqlHelper.IUD(list1)
        'sqlConnect.InsertData(sql.ToString())
    End Sub

    Private Sub InsertMonthData()
        Dim startTime As String
        Dim endTime As String
        Dim sql As New StringBuilder
        '首先判断是否已经有小时数据
        Dim sqlHelper As SQLiteHelper
        Dim list1 As New ArrayList

        sqlHelper = New SQLiteHelper(ConfigFilePath)

        '首先判断是否已经有小时数据
        startTime = Now.AddMonths(-1).ToString("yyyy-MM")
        endTime = Now.ToString("yyyy-MM") & "-01"

        sql.Append("select * from 月数据 where samplingtime ='" & startTime & "'")
        sqlHelper.SelectToDt(sql.ToString)
        If sqlHelper.RowsAffected > 0 Then Exit Sub

        startTime = startTime & "-01"
        sql.Clear()
        sql.Append("Insert into 月数据 (")
        sql.Append("samplingtime,ch4conc_h,ch4conc,ch4conc_z,ch4conc_p,nmhcconc_h,nmhcconc,nmhcconc_z,nmhcconc_p,tchconc_h,tchconc,tchconc_z,tchconc_p,")
        sql.Append("particulateconc_h,particulateconc,particulateconc_z,particulateconc_p,")
        sql.Append("o2conc_h,o2conc,tempconc,pressconc,speedconc,flowconc_g,flowconc_b,humiconc,fuheconc,REMARK )")
        sql.Append(" Select strftime('%Y-%m',samplingtime),")
        sql.Append(sqlDayString)
        sql.Append(" from 日数据")
        sql.Append(" where samplingtime >= '" & startTime & "' And samplingtime < '" & endTime & "'")
        sql.Append(" group by strftime('%Y.%m',samplingtime)")

        list1.Add(sql.ToString())
        sqlHelper.IUD(list1)
        'sqlConnect.InsertData(sql.ToString())
    End Sub

#End Region


#Region "污染源在线监测中的数据转换"
    ''' <summary>
    ''' 工况浓度转标况浓度
    ''' </summary>
    ''' <param name="value">测量参数的工况浓度</param>
    ''' <returns>测量参数的标况浓度</returns>
    Private Function WorkingToStandard(value As Single) As Single

        Return value * (101325 / (keyParameter.press + saveparameter.pressconc)) * ((273 + saveparameter.tempconc) / 273) * (1 / (1 - saveparameter.humiconc / 100))
    End Function

    ''' <summary>
    ''' 折算浓度计算,首先判断是否要折算，如果需要折算，还需要继续判断氧气浓度是否大于折算限值
    ''' </summary>
    ''' <returns>测量参数的折算氧系数</returns>
    Private Function ConvertConc() As Single
        Dim a As Single

        If keyParameter.corverFlag = 1 Then
            If o2data.O2Conc >= keyParameter.limitO2 Then
                a = 1
            Else
                a = 0.21 / ((0.21 - o2data.O2Conc / 100) * keyParameter.codffO2)
            End If
        Else
            a = 1
        End If

        Return a
    End Function

    ''' <summary>
    ''' 速度计算公式
    ''' </summary>
    ''' <returns>流速</returns>
    Private Function SpeedCalculate() As Single
        If saveparameter.speedconc >= 0 Then
            Return 0.076 * keyParameter.pitot * keyParameter.speed * Math.Sqrt((273 + saveparameter.tempconc) * saveparameter.speedconc)
        Else
            Return 0
        End If
    End Function

    ''' <summary>
    ''' 工况流量
    ''' </summary>
    ''' <returns>工况流量</returns>
    Private Function WorkingFlowCount() As Single
        Return 3600 * keyParameter.area * saveparameter.speedconc

    End Function

    '标况流量
    Private Function StandardFlowCount() As Single
        Return WorkingFlowCount() * (273 / (273 + saveparameter.tempconc)) * ((saveparameter.pressconc + keyParameter.press) / 101325) * (1 - saveparameter.humiconc / 100)
    End Function
#End Region


    ''' <summary>
    ''' 显示并计算所有的参数，由于采用的是热湿法测量甲烷、非甲烷总烃，所以测量得到的都是湿基浓度，需要进行很多转换
    ''' </summary>
    Private Sub DisPlayAllConcentration()

        Dim strbuf As String

        '采集的数据为湿基浓度，赋值给湿基浓度
        saveparameter.ch4conc_h = analyzerdata.CH4Conc                         '分析仪采集上来的甲烷浓度
        saveparameter.nmhcconc_h = analyzerdata.TCHConc - analyzerdata.CH4Conc '非甲烷总烃浓度（总烃 - 甲烷）
        saveparameter.tchconc_h = analyzerdata.TCHConc                         '分析仪采集上来的总烃浓度
        saveparameter.particulateconc_h = particulatedata.ParticulateConc '颗粒物采集仪获取的 颗粒物湿基浓度

        'saveparameter.ch4conc_h = ch4test
        'saveparameter.nmhcconc_h = nmhctest
        'saveparameter.tchconc_h = tchtest

        '氧气仪上读上来的是湿基氧气浓度
        saveparameter.o2conc_h = o2data.O2Conc
        saveparameter.humiconc = o2data.HumiConc     '此时湿度从氧气仪中读取

        saveparameter.o2conc = saveparameter.o2conc_h
        '赋值给标干浓度，通过下面的公式计算
        '临时注释，后面要使用
        saveparameter.ch4conc = Math.Round(saveparameter.ch4conc_h / (1 - saveparameter.humiconc / 100), 3)
        saveparameter.nmhcconc = Math.Round(saveparameter.nmhcconc_h / (1 - saveparameter.humiconc / 100), 3)
        saveparameter.tchconc = Math.Round(saveparameter.tchconc_h / (1 - saveparameter.humiconc / 100), 3)

        '根据氧气湿基浓度和湿度值计算氧气标干值
        saveparameter.o2conc = Math.Round(saveparameter.o2conc_h / (1 - saveparameter.humiconc / 100), 2)
        ' 计算颗粒物 干基浓度 add by wx 2024.08.20 16:13
        saveparameter.particulateconc = Math.Round(saveparameter.particulateconc_h / (1 - saveparameter.humiconc / 100), 3)


        Dim linshiA As String
        Dim linshiB As String


        analyzerdata.O2Value = Convert.ToUInt16(saveparameter.o2conc_h)

        'saveparameter.ch4conc = saveparameter.ch4conc_h
        'saveparameter.nmhcconc = saveparameter.nmhcconc_h
        'saveparameter.tchconc = saveparameter.tchconc_h

        '当使用旧版本的PLC时候使用
        saveparameter.tempconc = plcdata.TEMPConc
        saveparameter.pressconc = plcdata.PRESSConc
        ' saveparameter.humiconc = plcdata.HUMIDConc '此时湿度从PLC中读取
        saveparameter.speedconc = plcdata.SPEEDConc

        '当使用smartPLC使用使用
        'saveparameter.tempconc = plctcp.TEMPConc
        'saveparameter.pressconc = plctcp.PRESSConc
        'saveparameter.humiconc = plctcp.HUMIDConc
        'saveparameter.speedconc = plctcp.SPEEDConc
        '如果plcdata上报的动压信号，需要通过动压值计算流速
        '计算流速
        'saveparameter.speedconc = SpeedCalculate()

        '----------------------------------------------
        '预留校准，正式使用不能加上下面的校准
        '计算流速
        linshiA = ReadConfigIni("调试参数", "流速A", "1")
        linshiB = ReadConfigIni("调试参数", "流速B", "0")
        '直接测量为流速，不需要转换
        'saveparameter.speedconc = SpeedCalculate() * CSng(linshiA) + CSng(linshiB)
        '测量动压，转换为流速
        'saveparameter.speedconc = SpeedCalculate() * CSng(linshiA) + CSng(linshiB)

        '计算温度
        linshiA = ReadConfigIni("调试参数", "温度A", "1")
        linshiB = ReadConfigIni("调试参数", "温度B", "0")
        saveparameter.tempconc = saveparameter.tempconc * CSng(linshiA) + CSng(linshiB)

        '----------------------------------------------
        saveparameter.flowconc_g = WorkingFlowCount()
        saveparameter.flowconc_b = StandardFlowCount()
        '折算和排放量计算
        saveparameter.ch4conc_z = Math.Round(saveparameter.ch4conc * ConvertConc(), 3)
        saveparameter.nmhcconc_z = Math.Round(saveparameter.nmhcconc * ConvertConc(), 3)
        saveparameter.tchconc_z = Math.Round(saveparameter.tchconc * ConvertConc(), 3)
        '颗粒物 add by wx at 2024.08.20 16:23
        saveparameter.particulateconc_z = Math.Round(saveparameter.particulateconc * ConvertConc(), 3)

        saveparameter.ch4conc_p = Math.Round(saveparameter.ch4conc * saveparameter.flowconc_b * Math.Pow(10, -6), 2)
        saveparameter.nmhcconc_p = Math.Round(saveparameter.nmhcconc * saveparameter.flowconc_b * Math.Pow(10, -6), 2)
        saveparameter.tchconc_p = Math.Round(saveparameter.tchconc * saveparameter.flowconc_b * Math.Pow(10, -6), 2)
        '颗粒物 add by wx at 2024.08.20 16:23
        saveparameter.particulateconc_p = Math.Round(saveparameter.particulateconc * saveparameter.flowconc_b * Math.Pow(10, -6), 2)

        '===========显示在控件上================================
        strbuf = saveparameter.ch4conc.ToString("F2") & "," &
                 saveparameter.ch4conc_z.ToString("F2") & "," &
                 saveparameter.ch4conc_p.ToString("F2")
        '当大于设置的报警值，背景颜色变化
        Usercontrol_ch4.SetValue(strbuf)
        If saveparameter.ch4conc > alarmParameter.ch4alarmup Then
            If Usercontrol_ch4.DisplayColor <> Color.Gray Then
                Usercontrol_ch4.DisplayColor = Color.Gray
            End If
        Else
            If Usercontrol_ch4.DisplayColor <> Color.LightSeaGreen Then
                Usercontrol_ch4.DisplayColor = Color.LightSeaGreen
            End If
        End If

        strbuf = saveparameter.tchconc.ToString("F2") & "," &
                 saveparameter.tchconc_z.ToString("F2") & "," &
                 saveparameter.tchconc_p.ToString("F2")
        Usercontrol_Tch.SetValue(strbuf)

        If saveparameter.nmhcconc > alarmParameter.nmhcalarmup Then
            If Usercontrol_Nmhc.DisplayColor <> Color.Gray Then
                Usercontrol_Nmhc.DisplayColor = Color.Gray
            End If
        Else
            If Usercontrol_Nmhc.DisplayColor <> Color.LightSeaGreen Then
                Usercontrol_Nmhc.DisplayColor = Color.LightSeaGreen
            End If
        End If

        strbuf = saveparameter.nmhcconc.ToString("F2") & "," &
                 saveparameter.nmhcconc_z.ToString("F2") & "," &
                 saveparameter.nmhcconc_p.ToString("F2")
        Usercontrol_Nmhc.SetValue(strbuf)

        If saveparameter.tchconc > alarmParameter.tchalarmup Then
            If Usercontrol_Tch.DisplayColor <> Color.Gray Then
                Usercontrol_Tch.DisplayColor = Color.Gray
            End If
        Else
            If Usercontrol_Tch.DisplayColor <> Color.LightSeaGreen Then
                Usercontrol_Tch.DisplayColor = Color.LightSeaGreen
            End If
        End If

        '显示颗粒物 干基,折算，排放值 add by wx at 2024.08.20 16:29
        strbuf = saveparameter.particulateconc.ToString("F2") & "," &
                 saveparameter.particulateconc_z.ToString("F2") & "," &
                 saveparameter.particulateconc_p.ToString("F2")
        Usercontrol_Particulate.SetValue(strbuf)

        If saveparameter.particulateconc > alarmParameter.particulatealarmup Then
            If Usercontrol_Particulate.DisplayColor <> Color.Gray Then
                Usercontrol_Particulate.DisplayColor = Color.Gray
            End If
        Else
            If Usercontrol_Particulate.DisplayColor <> Color.LightSeaGreen Then
                Usercontrol_Particulate.DisplayColor = Color.LightSeaGreen
            End If
        End If

        strbuf = saveparameter.tempconc.ToString("F2")
        Usercontrol_TEMP.SetValue(strbuf)

        strbuf = saveparameter.pressconc.ToString("F1")
        Usercontrol_PRESS.SetValue(strbuf)

        strbuf = saveparameter.speedconc.ToString("F2") & "," &
                 saveparameter.flowconc_g.ToString("F0") & "," &
                 saveparameter.flowconc_b.ToString("F0")
        Usercontrol_SPEED.SetValue(strbuf)

        strbuf = saveparameter.humiconc.ToString("F2")
        Usercontrol_HUMI.SetValue(strbuf)

        strbuf = saveparameter.o2conc.ToString("F2")
        Usercontrol_O2.SetValue(strbuf)

        strbuf = saveparameter.ch4conc_h.ToString("F2")
        Usercontrol_ch4_h.SetValue(strbuf)

        strbuf = saveparameter.tchconc_h.ToString("F2")
        Usercontrol_tch_h.SetValue(strbuf)

        strbuf = saveparameter.nmhcconc_h.ToString("F2")
        Usercontrol_nmhc_h.SetValue(strbuf)

        strbuf = saveparameter.particulateconc_h.ToString("F2")
        Usercontrol_Particulate_h.SetValue(strbuf) '颗粒物 湿基 浓度 add by wx at 2024.08.20 16:36

        strbuf = saveparameter.o2conc_h.ToString("F2")
        Usercontrol_o2_h.SetValue(strbuf)

        Temperature1.SetAnalyzerValue(plcdata.TEMP_AI1.ToString())
        Temperature2.SetAnalyzerValue(plcdata.TEMP_AI2.ToString())
        Temperature3.SetAnalyzerValue(plcdata.TEMP_AI3.ToString())

        '每5秒的浓度值赋值给list，为计算每分钟平均值
        minuteAvg.Add(saveparameter)
    End Sub

    Private Sub MinuteAverage()
        Dim listbuff As New VOCParameterName
        Dim iIndex As Integer

        If minuteAvg.Count < 1 Then Exit Sub

        For iIndex = 0 To minuteAvg.Count - 1
            listbuff.ch4conc += minuteAvg.Item(iIndex).ch4conc
            listbuff.ch4conc_h += minuteAvg.Item(iIndex).ch4conc_h
            listbuff.ch4conc_p += minuteAvg.Item(iIndex).ch4conc_p
            listbuff.ch4conc_z += minuteAvg.Item(iIndex).ch4conc_z
            listbuff.nmhcconc += minuteAvg.Item(iIndex).nmhcconc
            listbuff.nmhcconc_h += minuteAvg.Item(iIndex).nmhcconc_h
            listbuff.nmhcconc_p += minuteAvg.Item(iIndex).nmhcconc_p
            listbuff.nmhcconc_z += minuteAvg.Item(iIndex).nmhcconc_z
            listbuff.tchconc += minuteAvg.Item(iIndex).tchconc
            listbuff.tchconc_h += minuteAvg.Item(iIndex).tchconc_h
            listbuff.tchconc_p += minuteAvg.Item(iIndex).tchconc_p
            listbuff.tchconc_z += minuteAvg.Item(iIndex).tchconc_z

            listbuff.particulateconc += minuteAvg.Item(iIndex).particulateconc
            listbuff.particulateconc_h += minuteAvg.Item(iIndex).particulateconc_h
            listbuff.particulateconc_p += minuteAvg.Item(iIndex).particulateconc_p
            listbuff.particulateconc_z += minuteAvg.Item(iIndex).particulateconc_z

            listbuff.o2conc += minuteAvg.Item(iIndex).o2conc
            listbuff.o2conc_h += minuteAvg.Item(iIndex).o2conc_h
            listbuff.tempconc += minuteAvg.Item(iIndex).tempconc
            listbuff.pressconc += minuteAvg.Item(iIndex).pressconc
            listbuff.speedconc += minuteAvg.Item(iIndex).speedconc
            listbuff.humiconc += minuteAvg.Item(iIndex).humiconc
            listbuff.flowconc_b += minuteAvg.Item(iIndex).flowconc_b
            listbuff.flowconc_g += minuteAvg.Item(iIndex).flowconc_g
            listbuff.fuheconc += minuteAvg.Item(iIndex).fuheconc
        Next

        listbuff.ch4conc = Math.Round(listbuff.ch4conc / minuteAvg.Count, 3)
        listbuff.ch4conc_h = Math.Round(listbuff.ch4conc_h / minuteAvg.Count, 3）
        listbuff.ch4conc_p = Math.Round(listbuff.ch4conc_p / minuteAvg.Count, 3）
        listbuff.ch4conc_z = Math.Round(listbuff.ch4conc_z / minuteAvg.Count, 3）
        listbuff.nmhcconc = Math.Round(listbuff.nmhcconc / minuteAvg.Count, 3）
        listbuff.nmhcconc_h = Math.Round(listbuff.nmhcconc_h / minuteAvg.Count, 3)
        listbuff.nmhcconc_p = Math.Round(listbuff.nmhcconc_p / minuteAvg.Count, 3)
        listbuff.nmhcconc_z = Math.Round(listbuff.nmhcconc_z / minuteAvg.Count, 3)
        listbuff.tchconc = Math.Round(listbuff.tchconc / minuteAvg.Count, 3)
        listbuff.tchconc_h = Math.Round(listbuff.tchconc_h / minuteAvg.Count, 3)
        listbuff.tchconc_p = Math.Round(listbuff.tchconc_p / minuteAvg.Count, 3)
        listbuff.tchconc_z = Math.Round(listbuff.tchconc_z / minuteAvg.Count, 3)
        '计算 颗粒物的  干基、湿基、排放、折算的 分钟平均值
        listbuff.particulateconc = Math.Round(listbuff.particulateconc / minuteAvg.Count, 3)
        listbuff.particulateconc_h = Math.Round(listbuff.particulateconc_h / minuteAvg.Count, 3)
        listbuff.particulateconc_p = Math.Round(listbuff.particulateconc_p / minuteAvg.Count, 3)
        listbuff.particulateconc_z = Math.Round(listbuff.particulateconc_z / minuteAvg.Count, 3)

        listbuff.o2conc = Math.Round(listbuff.o2conc / minuteAvg.Count, 2)
        listbuff.o2conc_h = Math.Round(listbuff.o2conc_h / minuteAvg.Count, 2)
        listbuff.tempconc = Math.Round(listbuff.tempconc / minuteAvg.Count, 2)
        listbuff.pressconc = Math.Round(listbuff.pressconc / minuteAvg.Count, 3)
        listbuff.speedconc = Math.Round(listbuff.speedconc / minuteAvg.Count, 2)
        listbuff.humiconc = Math.Round(listbuff.humiconc / minuteAvg.Count, 2)
        listbuff.flowconc_b = Math.Round(listbuff.flowconc_b / minuteAvg.Count, 1)
        listbuff.flowconc_g = Math.Round(listbuff.flowconc_g / minuteAvg.Count, 2)
        listbuff.fuheconc = Math.Round(listbuff.fuheconc / minuteAvg.Count)

        saveMinuteAverage = listbuff


        '如果当前系统按钮 是 运行时
        If calPlatformValve = "N" And calOrzerobuff = "N" Then
            saveMinuteAverage.note = "N"
        Else
            saveMinuteAverage.note = "C"
        End If

        '如果当前系统是维护时
        'saveMinuteAverage.note = "M"

    End Sub

    Private Sub DisplayAllParameter()
        'TextBox1.Text = analyzer.CH4Conc
        'TextBox2.Text = analyzer.TCHConc
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        Dim frm As New FormReport

        frm.ShowDialog()

    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        Dim frm As New FormPLCStatus
        frm.ShowDialog()
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        Dim frm As New FormLogView
        frm.ShowDialog()
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        Dim frm As New FormChangeLog
        frm.ShowDialog()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim frm As New FormUartSetting
        frm.StartPosition = FormStartPosition.CenterParent
        frm.ShowDialog()

    End Sub

    Private Sub FormMain_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Dim result As DialogResult
        result = MessageBox.Show("请确认是否退出系统?", "退出", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If DialogResult.Yes = result Then
            'Thread.CurrentThread.Abort()
            '记录退出时间
            WriteLogOutInfo()
            WriteLogToText("power", "power", "", False)
            Application.Exit()
        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        Static zeroOrcalCount As Integer

        RefreshChartData()

        '根据是否校准，判断标志位
        If CalOrZeroFlag <> 0 Then zeroOrcalCount += 1

        If CalOrZeroFlag = 1 AndAlso zeroOrcalCount >= calibrationPara.zeroTime Then
            CalOrZeroFlag = 0
            zeroOrcalCount = 0
        End If

        If CalOrZeroFlag = 2 AndAlso zeroOrcalCount >= calibrationPara.rangeTime Then
            CalOrZeroFlag = 0
            zeroOrcalCount = 0
        End If

        If CalOrZeroFlag = 1 Then calOrzerobuff = "C"
        If CalOrZeroFlag = 2 Then calOrzerobuff = "C"
        '暂时这样定义，还需要读取PLC状态
        If CalOrZeroFlag = 0 Then calOrzerobuff = "N"


        'If plctcp.PLCStatus.systemRun Then
        '    Label1.Text = "系统运行"
        '    If plctcp.PLCStatus.pitotFlow Then Label1.Text = "皮托管反吹"
        '    If plctcp.PLCStatus.gasFlow Then Label1.Text = "烟气反吹"
        'End If

        'If plctcp.PLCStatus.systemMain Then
        '    Label1.Text = "系统维护"
        'End If
    End Sub


    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Chart1.Series(0).Enabled = CheckBox1.Checked
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Chart1.Series(1).Enabled = CheckBox2.Checked
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        Chart1.Series(2).Enabled = CheckBox3.Checked
    End Sub
    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        Chart1.Series(3).Enabled = CheckBox4.Checked
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        Chart1.Series(4).Enabled = CheckBox5.Checked
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        Chart1.Series(5).Enabled = CheckBox6.Checked
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        Chart1.Series(6).Enabled = CheckBox7.Checked
    End Sub

    Private Sub RefreshChartData()

        'Dim i As Integer
        Dim rand1 As New Random
        Dim rand2 As New Random
        Dim rand3 As New Random
        '当数据点小于总数，
        listx.Add(Now.ToString("mm:ss"))
        list1.Add(saveparameter.ch4conc)
        list2.Add(saveparameter.tchconc)
        list3.Add(saveparameter.nmhcconc)
        list4.Add(saveparameter.o2conc)
        list5.Add(saveparameter.tempconc)
        list6.Add(saveparameter.pressconc)
        list7.Add(saveparameter.flowconc_b)

        If list1.Count > pointCount Then
            listx.RemoveAt(0)
            list1.RemoveAt(0)
            list2.RemoveAt(0)
            list3.RemoveAt(0)
            list4.RemoveAt(0)
            list5.RemoveAt(0)
            list6.RemoveAt(0)
            list7.RemoveAt(0)
        End If

        Chart1.Series(0).Points.DataBindXY(listx, list1)
        Chart1.Series(1).Points.DataBindXY(listx, list2)
        Chart1.Series(2).Points.DataBindXY(listx, list3)
        Chart1.Series(3).Points.DataBindXY(listx, list4)
        Chart1.Series(4).Points.DataBindXY(listx, list5)
        Chart1.Series(5).Points.DataBindXY(listx, list6)
        Chart1.Series(6).Points.DataBindXY(listx, list7)
    End Sub


    Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs) Handles ToolStripButton8.Click
        Dim frm As New FormChart
        frm.ShowDialog()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Dim frm As New FormKayParameter
        frm.ShowDialog()
    End Sub

    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        '用户登录对话框
        'Dim frm As New FormLogin
        'frm.ShowDialog()

        Using dialog As New FormLogin
            '订阅对话框的关闭事件
            If dialog.ShowDialog() = DialogResult.OK Then
                Task.Run(Sub()
                             '使用Invoke异步刷新界面
                             Me.Invoke(New Action(AddressOf RefreshFormMain))
                         End Sub)
            End If
        End Using
    End Sub

    Private Sub RefreshFormMain()
        '根据用户类型控制界面UI的使用权限
        '管理员开发所有界面入口
        If UserRole = EnumDefine.UserRoles.Inspector Then
            AdjustInspectorUI()
        ElseIf UserRole = EnumDefine.UserRoles.Administrator Then
            switchToAdministratorUI()
        Else

        End If
    End Sub

    ''' <summary>
    ''' 显示“测量数据”配置
    ''' </summary>
    Private Sub selectMeasuredData()
        '选择“测量数据”配置页面
        UiButton1.ForeColor = Color.Black
        UiButton2.ForeColor = Color.White
        UiButton3.ForeColor = Color.White
        UiButton4.ForeColor = Color.White

        '“测量温度”
        FlowLayoutPanel2.Visible = False
        '“实时曲线”
        SplitContainer2.Visible = False

        '“测量数据”
        FlowLayoutPanel1.Visible = True

    End Sub

    Private Sub AdjustInspectorUI()
        '菜单“设置”项灰显
        Me.设置SToolStripMenuItem.Visible = False

        '左侧导航栏
        '巡视员隐藏左侧导航栏的“系统操作”
        UiButton3.Visible = False
        'Panel1“系统操作”配置页面
        Panel1.Visible = False

        '工具栏
        '串口设置
        ToolStripButton1.Visible = False

        '参数设置
        ToolStripButton2.Visible = False

        '事件报告
        ToolStripButton2.Visible = False

        '切换为“巡视员”后默认选择“测量数据”配置页面
        selectMeasuredData()

    End Sub

    ''' <summary>
    ''' 由“监视员”切换回“管理员”
    ''' </summary>
    ''' 
    Private Sub switchToAdministratorUI()
        '菜单“设置”项灰显
        Me.设置SToolStripMenuItem.Visible = True

        '左侧导航栏
        UiButton3.Visible = True

        '工具栏
        '串口设置
        ToolStripButton1.Visible = True

        '参数设置
        ToolStripButton2.Visible = True

        '事件报告
        ToolStripButton2.Visible = True

        '切换为“巡视员”后默认选择“测量数据”配置页面
        selectMeasuredData()
    End Sub

    Private Sub ToolStripButton9_Click(sender As Object, e As EventArgs) Handles ToolStripButton9.Click
        Dim frm As New FormDataReport

        frm.ShowDialog()
    End Sub

    Private Sub SendMinuteData()
        Dim strbuf As New StringBuilder

        strbuf.Append(Now.ToString("yyyyMMddHHmmssfff") & ";")
        'strbuf.Append("a00000-Cou=" & Math.Round(saveMinuteAverage.flowconc_b, 1) & ";")
        strbuf.Append("a00000-Avg=" & Math.Round(saveMinuteAverage.flowconc_b, 1) & ";") '废气
        strbuf.Append("a19001-Avg=" & Math.Round(saveMinuteAverage.o2conc, 1) & ",") '氧气 平均值
        strbuf.Append("a19001-ZsAvg=" & Math.Round(saveMinuteAverage.o2conc_h, 1) & ";") '氧气 平均折算值
        strbuf.Append("a24088-Avg=" & Math.Round(saveMinuteAverage.nmhcconc, 3) & ",") '非甲烷总烃 平均值
        strbuf.Append("a24088-ZsAvg=" & Math.Round(saveMinuteAverage.nmhcconc_z, 3) & ";") '非甲烷总烃 平均折算值
        strbuf.Append("a24087-Avg=" & Math.Round(saveMinuteAverage.tchconc, 2) & ",") '碳氢化合物 平均值
        strbuf.Append("a24087-ZsAvg=" & Math.Round(saveMinuteAverage.tchconc_z, 2) & ";") '碳氢化合物 平均折算值
        strbuf.Append("a05002-Avg=" & Math.Round(saveMinuteAverage.ch4conc, 1) & ",") '甲烷 平均值
        strbuf.Append("a05002-ZsAvg=" & Math.Round(saveMinuteAverage.ch4conc_z, 1) & ";") '甲烷 平均折算值
        ' 颗粒物 add by wx at 2024.08.21 10:30
        strbuf.Append("a34013-Avg=" & Math.Round(saveMinuteAverage.particulateconc, 1) & ",") '颗粒物 平均值
        strbuf.Append("a34013-ZsAvg=" & Math.Round(saveMinuteAverage.particulateconc_z, 1) & ";") '颗粒物 平均折算值


        strbuf.Append("a01012-Avg=" & Math.Round(saveMinuteAverage.tempconc, 1) & ";") '烟气温度
        strbuf.Append("a01011-Avg=" & Math.Round(saveMinuteAverage.speedconc, 3) & ",") '烟气流速
        strbuf.Append("a01014-Avg=" & Math.Round(saveMinuteAverage.humiconc, 3) & ";") '烟气湿度
        strbuf.Append("a01013-Avg=" & Math.Round(saveMinuteAverage.pressconc, 2)) '烟气压力  

        Try
            datasamplinganalyzer.Write(senddatatoserver.MinuteData_2051(strbuf.ToString()))

            WriteLogToText(Now.ToString("yyyyMMdd"), "Operation", strbuf.ToString())
        Catch ex As Exception

        End Try

    End Sub

    Private Sub SendRealTimeData() '上传污染物实时数据  
        Dim strbuf As New StringBuilder
        ''作为测试使用
        'saveparameter.ch4conc = 20.89
        'saveparameter.ch4conc_z = 20.37

        'saveparameter.tempconc = 125.36
        'saveparameter.speedconc = 25.45
        'saveparameter.pressconc = 101136
        'saveparameter.humiconc = 21.58

        ''废气
        'saveparameter.flowconc_b = 111.12
        'saveparameter.o2conc = 22.61
        'saveparameter.o2conc_h = 21.91
        'saveparameter.nmhcconc = 0.508
        'saveparameter.nmhcconc_z = 0.492

        'saveparameter.tchconc = 1.186
        'saveparameter.tchconc_z = 1.186

        'saveparameter.particulateconc = 2.356
        'saveparameter.particulateconc_z = 2.357

        strbuf.Append(Now.ToString("yyyyMMddHHmmssfff") & ";")
        strbuf.Append("a00000-Rtd=" & Math.Round(saveparameter.flowconc_b, 1) & ";") '废气 污染物实时采样数据
        'strbuf.Append("a19001-Rtd=" & Math.Round(saveparameter.o2conc, 1) & ",") '氧气 污染物实时采样数据 标杆值
        '2025.11.05 修改上传精度为2位小数
        strbuf.Append($"a19001-Rtd={Math.Round(saveparameter.o2conc, 2):F2},")
        '2024.10.17化一环境新K37
        'strbuf.Append("a19001w-Rtd=" & Math.Round(saveparameter.o2conc_h, 1) & ";") '氧气 湿基值
        strbuf.Append($"a19001w-Rtd={Math.Round(saveparameter.o2conc_h, 2):F2};")

        strbuf.Append("a24088-Rtd=" & Math.Round(saveparameter.nmhcconc, 3) & ",")
        strbuf.Append("a24088-ZsRtd=" & Math.Round(saveparameter.nmhcconc_z, 3) & ";")
        '碳氢化合物
        strbuf.Append("a24087-Rtd=" & Math.Round(saveparameter.tchconc, 2) & ",")
        strbuf.Append("a24087-ZsRtd=" & Math.Round(saveparameter.tchconc_z, 2) & ";")
        strbuf.Append("a05002-Rtd=" & Math.Round(saveparameter.ch4conc, 1) & ",")
        strbuf.Append("a05002-ZsRtd=" & Math.Round(saveparameter.ch4conc_z, 1) & ";")
        ' 颗粒物 add by wx at 2024.08.21 10:30
        strbuf.Append("a34013-Rtd=" & Math.Round(saveparameter.particulateconc, 1) & ",") '颗粒物 平均值
        strbuf.Append("a34013-ZsRtd=" & Math.Round(saveparameter.particulateconc_z, 1) & ";") '颗粒物 平均折算值

        strbuf.Append("a01012-Rtd=" & Math.Round(saveparameter.tempconc, 1) & ";")
        strbuf.Append("a01011-Rtd=" & Math.Round(saveparameter.speedconc, 3) & ",")
        strbuf.Append("a01014-Rtd=" & Math.Round(saveparameter.humiconc, 3) & ";")
        strbuf.Append("a01013-Rtd=" & Math.Round(saveparameter.pressconc, 2))
        Try
            datasamplinganalyzer.Write(senddatatoserver.RealTimeData_2011(strbuf.ToString()))
        Catch ex As Exception

        End Try

    End Sub

    Private Sub SendHourData()
        Dim strbuf As New StringBuilder
        Dim sqlHelper As SQLiteHelper
        Dim strStartDate As String

        strStartDate = Now.AddHours(-1).ToString("yyyy-MM-dd HH") & ":00"
        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        Dim sql As String = "SELECT samplingtime,flowconc_b,o2conc,o2conc_h,
                            nmhcconc,nmhcconc_z,tchconc,tchconc_z,ch4conc,ch4conc_z,
                            particulateconc,particulateconc_z,
                            tempconc,speedconc,humiconc,pressconc FROM 小时数据 where samplingtime ='" & strStartDate & "' "

        sqlHelper.SelectToArray(sql)
        If sqlHelper.DatasArr.Length < 1 Then Exit Sub

        strbuf.Append(Convert.ToDateTime(sqlHelper.DatasArr(0, 0)).ToString("yyyyMMddHHmmssfff") & ";")
        'strbuf.Append("a00000-Cou=" & Math.Round(Convert.ToDouble(sqlHelper.DatasArr(0, 1)), 1) & ";")
        strbuf.Append("a00000-Avg=" & Math.Round(Convert.ToDouble(sqlHelper.DatasArr(0, 1)), 1) & ";")
        strbuf.Append("a19001-Avg=" & Math.Round(Convert.ToDouble(sqlHelper.DatasArr(0, 2)), 1) & ",")
        strbuf.Append("a19001-ZsAvg=" & Math.Round(Convert.ToDouble(sqlHelper.DatasArr(0, 3)), 1) & ";")
        strbuf.Append("a24088-Avg=" & Math.Round(Convert.ToDouble(sqlHelper.DatasArr(0, 4)), 3) & ",")
        strbuf.Append("a24088-ZsAvg=" & Math.Round(Convert.ToDouble(sqlHelper.DatasArr(0, 5)), 3) & ";")
        strbuf.Append("a24087-Avg=" & Math.Round(Convert.ToDouble(sqlHelper.DatasArr(0, 6)), 2) & ",")
        strbuf.Append("a24087-ZsAvg=" & Math.Round(Convert.ToDouble(sqlHelper.DatasArr(0, 7)), 2) & ";")
        strbuf.Append("a05002-Avg=" & Math.Round(Convert.ToDouble(sqlHelper.DatasArr(0, 8)), 1) & ",")
        strbuf.Append("a05002-ZsAvg=" & Math.Round(Convert.ToDouble(sqlHelper.DatasArr(0, 9)), 1) & ";")
        ' 颗粒物  add by wx at 2024.08.21 16:00
        strbuf.Append("a34013-Avg=" & Math.Round(Convert.ToDouble(sqlHelper.DatasArr(0, 10)), 1) & ",")
        strbuf.Append("a34013-ZsAvg=" & Math.Round(Convert.ToDouble(sqlHelper.DatasArr(0, 11)), 1) & ";")

        strbuf.Append("a01012-Avg=" & Math.Round(Convert.ToDouble(sqlHelper.DatasArr(0, 12)), 1) & ";")
        strbuf.Append("a01011-Avg=" & Math.Round(Convert.ToDouble(sqlHelper.DatasArr(0, 13)), 3) & ",")
        strbuf.Append("a01014-Avg=" & Math.Round(Convert.ToDouble(sqlHelper.DatasArr(0, 14)), 3) & ";")
        strbuf.Append("a01013-Avg=" & Math.Round(Convert.ToDouble(sqlHelper.DatasArr(0, 15)), 2))

        Try
            datasamplinganalyzer.Write(senddatatoserver.HourData_2061(strbuf.ToString()))
        Catch ex As Exception

        End Try


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim k As New ModbusConvert

        'Dim a As UShort() = {k.Int16ToUshort(-2), 5, 5}
        'Dim b As Single() = {0, 5, 5}


        'b = FloatToUshort(300)
        'Array.Copy(b, 0, a, 0, 2)
        'b = FloatToUshort(0)
        'Array.Copy(b, 0, a, 2, 2)
        'b = FloatToUshort(200)
        'Array.Copy(b, 0, a, 4, 2)
        'b = FloatToUshort(0)
        'Array.Copy(b, 0, a, 6, 2)

        'a = ArraySingleToUshort(b)

        'plctcp.setholdflag = True
        'plctcp.holdregadd = 100
        'plctcp.holdregdata = a

        InsertHourData()
    End Sub

    Private Sub UiButton1_Click(sender As Object, e As EventArgs) Handles UiButton1.Click
        FlowLayoutPanel1.Visible = True
        FlowLayoutPanel2.Visible = False
        Panel1.Visible = False
        SplitContainer2.Visible = False


        UiButton1.ForeColor = Color.Black
        UiButton2.ForeColor = Color.White
        UiButton3.ForeColor = Color.White
        UiButton4.ForeColor = Color.White
    End Sub

    Private Sub UiButton2_Click(sender As Object, e As EventArgs) Handles UiButton2.Click
        FlowLayoutPanel1.Visible = False
        FlowLayoutPanel2.Visible = True
        Panel1.Visible = False
        SplitContainer2.Visible = False

        UiButton1.ForeColor = Color.White
        UiButton2.ForeColor = Color.Black
        UiButton3.ForeColor = Color.White
        UiButton4.ForeColor = Color.White
    End Sub

    Private Sub UiButton4_Click(sender As Object, e As EventArgs) Handles UiButton4.Click
        FlowLayoutPanel1.Visible = False
        FlowLayoutPanel2.Visible = False
        Panel1.Visible = False
        SplitContainer2.Visible = True

        UiButton1.ForeColor = Color.White
        UiButton2.ForeColor = Color.White
        UiButton3.ForeColor = Color.White
        UiButton4.ForeColor = Color.Black
    End Sub

    Private Sub ToolStripButton10_Click(sender As Object, e As EventArgs) Handles ToolStripButton10.Click
        Dim frm As New Formformula
        frm.ShowDialog()
    End Sub

    Private Sub UiButton3_Click(sender As Object, e As EventArgs) Handles UiButton3.Click
        '“测量数据”
        FlowLayoutPanel1.Visible = False

        '“测量温度”界面暂时已不使用
        FlowLayoutPanel2.Visible = False

        '“系统操作”
        Panel1.Visible = True
        UiButton3.ForeColor = Color.Black

        '“实时曲线”
        SplitContainer2.Visible = False

        UiButton1.ForeColor = Color.White
        UiButton2.ForeColor = Color.White
        UiButton4.ForeColor = Color.White
    End Sub


    Private Sub InsertAdmin()
        Dim sqlHelper As SQLiteHelper
        Dim sql As New ArrayList
        Dim buf As String

        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        buf = "select * from userinfo where username='Admin'"
        sqlHelper.SelectToDt(buf)

        If sqlHelper.DatasDt.Rows.Count = 1 Then Exit Sub

        sql.Add("insert into UserInfo(UserName,userRight,userPW,Addtime) values ('Admin','管理员','123456','" & Now.ToString("yyyy-MM-dd HH:mm") & "')")

        sqlHelper.IUD(sql)
    End Sub

    Private Sub UiSwitch1_ValueChanged(sender As Object, value As Boolean) Handles UiSwitch1.ValueChanged

        Dim strbuf As String
        Dim buff As UInt16()
        ReDim buff(0)

        '如果“校标”已经打开了，则不能进行“校零”
        If CalOrZeroFlag <> 0 Then

            '会再次触发UiSwitch1的ValueChange
            UiSwitch1.Active = False
            Exit Sub
        End If

        If value Then
            CalOrZeroFlag = 1
            '校零使能之后，按钮不能直接执行复位而是通过点击“停止校准”按钮进行复位
            UiSwitch1.Enabled = False
            UiSwitch1.DisabledColor = UiSwitch1.ActiveColor

            strbuf = "分析仪开始校零，甲烷浓度=" & saveMinuteAverage.ch4conc_h & "，非甲烷总烃浓度=" & saveMinuteAverage.nmhcconc_h & "，氧气浓度=" & saveMinuteAverage.o2conc_h
            WriteCalLog(strbuf)


            buff(0) = 1

            plcdata.writeCom = True
            plcdata.writeAnaly = 1
            plcdata.writeAdd = 155
            plcdata.writeData = buff

            '防止用户没有执行点击“停止校准”按钮导致设备一直进行“校零”，定时自动执行校零复位
            '启动定时器
            If Not Timer5.Enabled Then
                Timer5.Interval = calibrationPara.calibrationTime * 1000
                Timer5.Start()
            End If

        Else
            '向串口下发停止“校零”
            buff(0) = 0

            plcdata.writeCom = True
            plcdata.writeAnaly = 1
            plcdata.writeAdd = 155
            plcdata.writeData = buff

            '定时器如果没有关闭，则关闭定时器
            If Timer5.Enabled Then
                Timer5.Stop()
            End If
        End If


    End Sub

    Private Sub UiButton5_Click(sender As Object, e As EventArgs) Handles UiButton5.Click
        '需要注意的是，点击“停止校准”按钮，只改变了界面的UI状态
        '返回正常采样状态
        CalOrZeroFlag = 0

        '触发UiSwitch1的ValueChanged
        UiSwitch1.Active = False
        'UiSwitch1能够响应用户点击
        UiSwitch1.Enabled = True

        '触发UiSwitch2的ValueChanged
        UiSwitch2.Active = False
        'UiSwitch2能够响应用户点击
        UiSwitch2.Enabled = True


        Dim strbuf As String

        strbuf = "分析仪校准完成，甲烷浓度=" & saveMinuteAverage.ch4conc_h & "，非甲烷总烃浓度=" & saveMinuteAverage.nmhcconc_h & "，氧气浓度=" & saveMinuteAverage.o2conc_h
        WriteCalLog(strbuf)
    End Sub

    Private Sub UiSwitch2_ValueChanged(sender As Object, value As Boolean) Handles UiSwitch2.ValueChanged

        Dim strbuf As String
        Dim buff As UInt16()
        ReDim buff(0)

        '如果校零已经打开了，则不能进行“校标”
        If CalOrZeroFlag <> 0 Then
            '“校标按钮”置为去激活状态
            UiSwitch2.Active = False
            Exit Sub
        End If

        If value Then
            CalOrZeroFlag = 2
            UiSwitch2.Enabled = False
            UiSwitch2.DisabledColor = UiSwitch1.ActiveColor

            strbuf = "分析仪开始校标，甲烷浓度=" & saveMinuteAverage.ch4conc_h & "，非甲烷总烃浓度=" & saveMinuteAverage.nmhcconc_h & "，氧气浓度=" & saveMinuteAverage.o2conc_h
            WriteCalLog(strbuf)

            buff(0) = 1

            plcdata.writeCom = True
            plcdata.writeAnaly = 1
            plcdata.writeAdd = 156
            plcdata.writeData = buff

            '防止用户没有执行点击“停止校准”按钮导致设置一直进行“校标”，定时自动执行校标复位
            '启动定时器
            If Not Timer6.Enabled Then
                Timer6.Interval = calibrationPara.rangeResetTime * 1000
                Timer6.Start()
            End If

        Else
            '停止“校标”
            buff(0) = 0

            plcdata.writeCom = True
            plcdata.writeAnaly = 1
            plcdata.writeAdd = 156
            plcdata.writeData = buff

            '定时器如果没有关闭，则关闭定时器
            If Timer6.Enabled Then
                Timer6.Stop()
            End If
        End If

    End Sub

    Private Sub 分钟数据ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 分钟数据ToolStripMenuItem.Click
        Dim frm As New FormMinuteReport
        frm.ShowDialog()

        frm.StartPosition = FormStartPosition.CenterParent

    End Sub

    Private Sub 登录LToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 登录LToolStripMenuItem.Click
        Dim frm As New FormLogin
        frm.ShowDialog()
        frm.StartPosition = FormStartPosition.CenterParent
    End Sub
    Private Sub 设置SToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 设置SToolStripMenuItem.Click
        Dim frm As New FormKayParameter
        frm.ShowDialog()
        frm.StartPosition = FormStartPosition.CenterParent
    End Sub



    Private Sub UiSwitch3_ValueChanged(sender As Object, value As Boolean) Handles UiSwitch3.ValueChanged
        Dim buff As UInt16()
        ReDim buff(0)

        If value Then
            '平台标准阀打开
            buff(0) = 1
            calPlatformValve = "C"
            plcdata.writeCom = True
            plcdata.writeAnaly = 1
            plcdata.writeAdd = 157
            plcdata.writeData = buff

            '启动定时器
            If Not Timer7.Enabled Then
                Timer7.Interval = calibrationPara.platformResetTime * 1000
                Timer7.Start()
            End If

        Else
            '平台标准阀关闭
            buff(0) = 0
            calPlatformValve = "N"
            plcdata.writeCom = True
            plcdata.writeAnaly = 1
            plcdata.writeAdd = 157
            plcdata.writeData = buff

            '定时器如果没有关闭，则关闭定时器
            If Timer7.Enabled Then
                Timer7.Stop()
            End If

        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Timer5_Tick(sender As Object, e As EventArgs) Handles Timer5.Tick
        '如果“校零”或者“校标”处于开启状态
        '定时器执行“停止校准”防止开关过热烧毁
        If CalOrZeroFlag = 1 Then
            'CalOrZeroFlag复位
            '返回正常采样状态
            CalOrZeroFlag = 0

            '执行停止“校零”
            UiSwitch1.Active = False
            UiSwitch1.Enabled = True
        End If

        '关闭当前定时器
        If Timer5.Enabled Then
            Timer5.Stop()
        End If
    End Sub

    Private Sub Timer6_Tick(sender As Object, e As EventArgs) Handles Timer6.Tick
        '如果“校标”处于开启状态
        '定时器执行“停止校标”防止开关过热烧毁
        If CalOrZeroFlag = 2 Then
            'CalOrZeroFlag复位
            '返回正常采样状态
            CalOrZeroFlag = 0

            '执行停止“校标”
            UiSwitch2.Active = False
            UiSwitch2.Enabled = True
        End If

        '关闭当前定时器
        If Timer6.Enabled Then
            Timer6.Stop()
        End If
    End Sub

    Private Sub Timer7_Tick(sender As Object, e As EventArgs) Handles Timer7.Tick
        '如果“平台阀”处于开启状态
        If String.Equals(calPlatformValve, "C") Then
            UiSwitch3.Active = False
            UiSwitch3.Enabled = True
        End If

        '在回调中关闭当前定时器
        '定时器也有可能在UiSwitch3的ValueChanged中关闭，先判断
        If Timer7.Enabled Then
            Timer7.Stop()
        End If
    End Sub
End Class
