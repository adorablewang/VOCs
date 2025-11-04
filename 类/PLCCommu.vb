Imports Modbus.Device
Imports Modbus.Utility
Imports Modbus.Data
Imports Modbus.Message
Imports System.Threading
Imports System.IO.Ports

Public Class PLCCommu

    Private master As ModbusMaster
    Private WithEvents serial As New SerialPort

    Private thread As Thread

    Private _tempConc As Single
    Private _pressConc As Single
    Private _speedConc As Single
    Private _humidConc As Single
    '差压
    Private _diffPressure As Single

    Private _temp1 As Int16
    Private _temp2 As Int16
    Private _temp3 As Int16
    '系统运行状态
    Private _systemRunStatus As Int16

    '写指令
    Public writeCom As Boolean
    '写地址
    Public writeAdd As UShort
    '写数字
    Public writeData As UShort()
    '选择表
    Public writeAnaly As UShort

    ' 发送质量计数
    Public sendCount As Integer

    'plc的IO输入
    Public inputIO(15) As Boolean
    'plc的ID输出
    Public ouputIO(15) As Boolean
    'plc的模拟量输入
    Public inputAI As Int16() = {0, 0, 0, 0, 0, 0, 0, 0}
    '是否写IO
    Public writeIOFlag As Boolean

    Private _serialConfig As String = "COM2,9600,N,8,1"
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

    '获取差压值
    Public ReadOnly Property DiffPressure() As Single
        Get
            Return _diffPressure
        End Get

    End Property

    Public ReadOnly Property HUMIDConc() As Single
        Get
            Return _humidConc
        End Get

    End Property

    Public ReadOnly Property TEMP_AI1() As Single
        Get
            Return _temp1 / 10
        End Get

    End Property

    Public ReadOnly Property TEMP_AI2() As Single
        Get
            Return _temp2 / 10
        End Get

    End Property

    Public ReadOnly Property TEMP_AI3() As Single
        Get
            Return _temp3 / 10
        End Get

    End Property

    Public Sub InitializePara()

        SetAnalyzerSerial()

        Try
            serial.Open()
        Catch ex As Exception
            Exit Sub
        End Try

        master = ModbusSerialMaster.CreateRtu(serial)

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


    ''' <summary>
    ''' 配置串口设置，端口、波特率、数据位、停止位
    ''' </summary>
    Private Sub SetAnalyzerSerial()
        Dim sTokens() As String

        serial = New SerialPort

        sTokens = _serialConfig.Split(",")
        serial.PortName = sTokens(0)
        serial.BaudRate = sTokens(1)
        If sTokens(2).ToUpper = "N" Then
            serial.Parity = Parity.None
        ElseIf sTokens(2).ToUpper = "O" Then
            serial.Parity = Parity.Odd
        ElseIf sTokens(2).ToUpper = "E" Then
            serial.Parity = Parity.Even
        End If

        serial.DataBits = sTokens(3)

        If sTokens(4).ToUpper = "1" Then
            serial.StopBits = StopBits.One
        ElseIf sTokens(4).ToUpper = "1.5" Then
            serial.StopBits = StopBits.OnePointFive
        ElseIf sTokens(4).ToUpper = "2" Then
            serial.StopBits = StopBits.Two
        End If

        'If sTokens(5) = 1 Then
        '    Try
        '        serial.Open()
        '    Catch ex As Exception

        '    End Try
        'End If

    End Sub

    Private Sub BackgroundProcess()
        Static i As Integer = 1

        Do While True
            If Not writeIOFlag Then
                Select Case i
                    Case 0
                        ReadConc()
                        i += 1
                    Case 1
                        'ReadInputAI()
                        i += 1
                    Case 2
                        'ReadInputIO()
                        i += 1
                    Case 3
                        'ReadOutputIO()
                        i = 0
                    Case Else
                        i = 0
                End Select

            Else

            End If

            Thread.Sleep(300)
        Loop
    End Sub

    '读保持寄存器
    Public Function ReadConc() As Boolean
        Dim a() As UShort
        Dim b(2) As Single

        If writeCom Then
            Try
                '参数设置中“量程参数”和系统操作中的“校零”、“校标”、“平台校准阀”
                '2024.10.15新增反吹
                '写入PLCCommu
                master.WriteMultipleRegisters(writeAnaly, writeAdd, writeData)
            Catch ex As Exception
                writeCom = False
            End Try
            writeCom = False
            'Exit Function
            Return True
        End If

        Try
            '通过PLCCommu读取
            a = master.ReadHoldingRegisters(1, 0, 11)
        Catch ex As Exception
            Return False
        End Try

        _tempConc = ModbusUtility.GetSingle(a(0), a(1))
        _pressConc = ModbusUtility.GetSingle(a(2), a(3))
        _speedConc = ModbusUtility.GetSingle(a(4), a(5))
        _humidConc = ModbusUtility.GetSingle(a(6), a(7))



        If _tempConc < 0 Then _tempConc = 0
        If _speedConc < 0 Then _speedConc = 0
        If _humidConc < 0 Then _humidConc = 0

        _temp1 = a(8)
        _temp2 = a(9)
        _temp3 = a(10)


        '获取硬件系统运行状态  ，运行 还是 维护
        Try
            a = master.ReadHoldingRegisters(1, 105, 1)
        Catch ex As Exception
            Return False
        End Try

        Dim systemStatus = a(0)

        Dim b_2 = systemStatus And (&H4) '运行状态
        Dim b_3 = systemStatus And (&H8) '维护状态
        Dim b_4 = systemStatus And (&H10) '平台校准状态



        Return True
    End Function

    ''' <summary>
    ''' 防爆温压流数据采集
    ''' </summary>
    ''' <returns></returns>
    Public Function ReadConcForKangce() As Boolean
        Dim a() As UShort

        If writeCom Then
            Try
                master.WriteMultipleRegisters(writeAnaly, writeAdd, writeData)
            Catch ex As Exception
                writeCom = False
            End Try
            writeCom = False
            'Exit Function
            Return True
        End If

        Try
            '需要根据PLC提供的寄存器地址（读取温压流上报的采集数据）
            '暂时从0地址读取
            '读取保持寄存器
            a = master.ReadHoldingRegisters(1, 0, 8)

            '根据用户手册依次为流速、静压、温度、差压（界面不显示）
            '流速
            _speedConc = ModbusUtility.GetSingle(a(0), a(1))
            '静压
            _pressConc = ModbusUtility.GetSingle(a(2), a(3))
            '温度
            _tempConc = ModbusUtility.GetSingle(a(4), a(5))
            '差压（需要确定是否有此数据）
            _diffPressure = ModbusUtility.GetSingle(a(6), a(7))

        Catch ex As Exception
            Return False
        End Try

        If _tempConc < 0 Then _tempConc = 0
        If _speedConc < 0 Then _speedConc = 0
        If _diffPressure < 0 Then _diffPressure = 0

        Return True
    End Function

    ''' <summary>
    ''' 读模拟量输入
    ''' </summary>
    ''' <returns></returns>
    Public Function ReadInputAI() As Boolean
        Dim a() As UShort
        Dim b(2) As Single

        Try
            a = master.ReadInputRegisters(1, 0, 16)
        Catch ex As Exception
            Return False
        End Try

        inputAI(0) = ModbusUtility.GetSingle(a(0), a(1))
        inputAI(1) = ModbusUtility.GetSingle(a(2), a(3))
        inputAI(2) = ModbusUtility.GetSingle(a(4), a(5))
        inputAI(3) = ModbusUtility.GetSingle(a(6), a(7))

        inputAI(4) = ModbusUtility.GetSingle(a(8), a(9))
        inputAI(5) = ModbusUtility.GetSingle(a(10), a(10))
        inputAI(6) = ModbusUtility.GetSingle(a(12), a(12))
        inputAI(7) = ModbusUtility.GetSingle(a(14), a(14))

        Return True
    End Function


    Public Function ReadInputIO() As Boolean

        Try
            inputIO = master.ReadInputs(1, 0, 16)
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    Public Function ReadOutputIO() As Boolean

        Try
            ouputIO = master.ReadCoils(1, 0, 16)
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    'Public Function WriteSingleCol() As Boolean

    '    'master.WriteSingleCoil(1,)
    'End Function
End Class
