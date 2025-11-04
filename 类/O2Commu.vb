Imports Modbus.Device
Imports Modbus.Utility
Imports Modbus.Data
Imports Modbus.Message
Imports System.Threading
Imports System.IO.Ports

Public Class O2COMMU

    Private master As ModbusMaster
    Private WithEvents serial As New SerialPort

    Private thread As Thread
    '氧气测量值
    Private _o2Conc As Single = 0
    '湿度测量值
    Private _humi As Single = 0
    Private _serialConfig As String = "COM3,9600,N,8,1"

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


    Public Property SerialConfig() As String
        Get
            Return _serialConfig
        End Get
        Set(ByVal value As String)
            _serialConfig = value
        End Set
    End Property

    Public ReadOnly Property O2Conc() As Single
        Get
            Return _o2Conc
        End Get

    End Property

    Public ReadOnly Property HumiConc() As Single
        Get
            Return _humi
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
        Dim i As Integer = 1

        Do While True
            ReadConc()
            Thread.Sleep(1000)
        Loop
    End Sub

    Public Function ReadConc() As Boolean
        Dim a() As UShort
        Dim b(2) As Single


        If writeCom Then
            Try
                master.WriteMultipleRegisters(writeAnaly, writeAdd, writeData)
            Catch ex As Exception
                writeCom = False
            End Try
            writeCom = False
            Return True
            Exit Function
        End If

        Try
            'a = master.ReadHoldingRegisters(1, 2, 2)
            '_o2Conc = Math.Round((ModbusUtility.GetUInt32(a(0), a(1)) / 1000000), 2)

            a = master.ReadInputRegisters(1, 0, 8)
            _o2Conc = Math.Round((ModbusUtility.GetSingle(a(2), a(3)) / 10000), 2)
            If _o2Conc < 0 Then _o2Conc = 0

            _humi = Math.Round((ModbusUtility.GetSingle(a(6), a(7)) / 10000), 2)
            If _humi < 0 Then _humi = 0
        Catch ex As Exception
            Return False
        End Try

        'Dim readholdr As ReadHoldingInputRegistersRequest
        'readholdr = New ReadHoldingInputRegistersRequest(&H3, &H1, &H0, &H4)

        'Dim readholdRes
        'Try
        '    readholdRes = master.ExecuteCustomMessage(Of ReadHoldingInputRegistersResponse)(readholdr)
        'Catch ex As Exception
        '    Return -1
        'End Try

        Return True
    End Function
End Class

