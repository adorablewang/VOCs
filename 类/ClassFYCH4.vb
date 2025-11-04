Imports Modbus.Device
Imports Modbus.Utility
Imports System.IO.Ports
Imports System.Threading

Public Class ClassFYCH4

    Private master As ModbusMaster
    Private WithEvents serial As New SerialPort

    Private thread As Thread

    Private _ch4Conc As Single
    Private _tchConc As Single

    Private _serialConfig As String = "COM1,9600,N,8,1"

    Private _o2value As UShort = 0
    Private oldo2value As UShort = 0
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

    Public receivedatavalue As String
    Public receivedatavalue1 As String
    Public WriteOnly Property O2Value() As UShort

        Set(ByVal value As UShort)
            _o2value = value
        End Set
    End Property


    Public Property SerialConfig() As String
        Get
            Return _serialConfig
        End Get
        Set(ByVal value As String)
            _serialConfig = value
        End Set
    End Property

    Public ReadOnly Property CH4Conc() As Single
        Get
            Return _ch4Conc
        End Get

    End Property

    Public ReadOnly Property TCHConc() As Single
        Get
            Return _tchConc
        End Get

    End Property

    Public Sub InitializePara()

        SetAnalyzerSerial()

        Try
            serial.Open()
        Catch ex As Exception

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
            thread.Sleep(1000)
        Loop
    End Sub

    Public Function ReadConc() As Boolean
        Dim a() As UShort
        Dim b(2) As Single

        'If writeCom Then
        '    Try
        '        master.WriteMultipleRegisters(writeAnaly, writeAdd, writeData)
        '    Catch ex As Exception
        '        writeCom = False
        '    End Try
        '    writeCom = False
        '    Exit Function
        'End If

        Select Case sendCount
            Case 1
                Try
                    '速越仪器
                    a = master.ReadInputRegisters(1, 0, 2)
                    _ch4Conc = a(0) / 100
                    _tchConc = a(1) / 100

                    WriteLogToText(Now.ToString("yyyyMMdd"), "communication", _tchConc & "," & _ch4Conc)
                Catch ex As Exception
                    receivedatavalue1 = Now.ToString("ss")
                    sendCount += 1
                    Return False
                End Try
            Case 2
                '需要把采集到的氧气浓度发送给分析仪，用来应对氧干扰
                'Try
                '    sendCount = 0
                '    If oldo2value <> _o2value Then
                '        oldo2value = _o2value
                '        master.WriteSingleRegister(10, 30, _o2value)
                '        WriteLogToText(Now.ToString("yyyyMMdd"), "communication", "氧气浓度:" & _o2value.ToString())
                '    End If
                'Catch ex As Exception
                '    WriteLogToText(Now.ToString("yyyyMMdd"), "communication", "写入错误," & ex.Message & "氧气浓度:" & _o2value.ToString())
                '    Return False
                'End Try
            Case Else
                sendCount = 0
        End Select

        sendCount += 1
        Return True
    End Function
End Class
