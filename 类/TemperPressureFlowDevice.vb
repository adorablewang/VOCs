Imports Modbus.Device
Imports Modbus.Utility
Imports Sunny.UI
Imports System.IO.Ports
Imports System.Threading

Public Class TemperPressureFlowDevice
    Private master As ModbusMaster
    Private WithEvents SerialComPort As SerialPort

    Private thread As Thread
    Private _serialConfig As String = "COM6,19200,E,8,1"

    '温度
    Private _tempConc As Single
    '压力
    Private _pressConc As Single
    '流速
    Private _speedConc As Single

    '通信参数默认地址
    Private Const DEV_ADDRESS As Integer = &HA
    Private Const REG_ASSRESS As Integer = &H7535

    '发送计数
    Public sendCount As Integer

    '记录异常时间
    Public exceptionMoment As String

    Public Property SerialConfig() As String
        Get
            Return _serialConfig
        End Get
        Set(ByVal value As String)
            _serialConfig = value
        End Set
    End Property

    Public ReadOnly Property TEMPConc() As Single
        Get
            Return _tempConc
        End Get

    End Property

    Public ReadOnly Property PRESSConc() As Single
        Get
            Return _pressConc
        End Get

    End Property

    Public ReadOnly Property SPEEDConc() As Single
        Get
            Return _speedConc
        End Get

    End Property

    ''' <summary>
    ''' 配置串口设置，端口、波特率、数据位、停止位
    ''' </summary>
    Private Sub SetAnalyzerSerial()
        Dim sTokens() As String

        SerialComPort = New SerialPort

        sTokens = _serialConfig.Split(",")
        SerialComPort.PortName = sTokens(0)
        SerialComPort.BaudRate = sTokens(1)

        Dim parityflag As String = sTokens(2)
        If parityflag.ToUpper = "N" Then
            SerialComPort.Parity = Parity.None
        ElseIf parityflag.ToUpper = "O" Then
            SerialComPort.Parity = Parity.Odd
        ElseIf parityflag.ToUpper = "E" Then
            SerialComPort.Parity = Parity.Even
        End If

        SerialComPort.DataBits = sTokens(3)

        If sTokens(4).ToUpper = "1" Then
            SerialComPort.StopBits = StopBits.One
        ElseIf sTokens(4).ToUpper = "1.5" Then
            SerialComPort.StopBits = StopBits.OnePointFive
        ElseIf sTokens(4).ToUpper = "2" Then
            SerialComPort.StopBits = StopBits.Two
        End If

    End Sub

    Public Sub InitializePara()

        SetAnalyzerSerial()

        Try
            SerialComPort.Open()
        Catch ex As Exception
            'UIMessageTip.ShowError("温压流检测仪串口打开失败！")

        End Try

        master = ModbusSerialMaster.CreateRtu(SerialComPort)

        With master
            .Transport.ReadTimeout = 200
            .Transport.WriteTimeout = 200
            .Transport.WaitToRetryMilliseconds = 200
            .Transport.Retries = 2
        End With

        thread = New Thread(AddressOf BackgroundProcess)
        thread.Start()
        thread.IsBackground = True
    End Sub

    Private Sub BackgroundProcess()

        Do While True
            ReadConc()
            Thread.Sleep(1000)
        Loop
    End Sub

    Public Function ReadConc() As Boolean
        Dim a() As UShort
        Select Case sendCount
            Case 1
                Try
                    '根据用户手册_温压流监测仪_C6读取输入寄存器(0x04)，起始地址30005（0x7535）
                    '读6个寄存器，参数分别是流速、压力、温度
                    a = master.ReadInputRegisters(DEV_ADDRESS, REG_ASSRESS, 6)

                    '流速
                    _speedConc = ModbusUtility.GetSingle(a(0), a(1))
                    '压力 单位Pa
                    _pressConc = ModbusUtility.GetSingle(a(2), a(3))
                    '温度
                    _tempConc = ModbusUtility.GetSingle(a(4), a(5))

                    If _tempConc < 0 Then _tempConc = 0
                    If _speedConc < 0 Then _speedConc = 0

                Catch ex As Exception
                    exceptionMoment = Now.ToString("yyyy-MM-dd HH:mm:ss")

                    Return False
                End Try

            Case Else
                sendCount = 1
        End Select

        Return True
    End Function



End Class
