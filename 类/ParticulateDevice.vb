Imports Modbus.Device
Imports Modbus.Utility
Imports Modbus.Data
Imports Modbus.Message
Imports System.Threading
Imports System.IO.Ports



Public Class ParticulateDevice

    Private master As ModbusMaster
    Private WithEvents serial As New SerialPort

    Private thread As Thread

    Private _particulateConc As Single  'add by wx 2024.08.19 15:42

    Private _serialConfig As String = "COM11,9600,N,8,1"

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



    Public Property SerialConfig() As String
        Get
            Return _serialConfig
        End Get
        Set(ByVal value As String)
            _serialConfig = value
        End Set
    End Property



    Public ReadOnly Property ParticulateConc() As Single 'add by wx 2024.08.19 15:42
        Get
            Return _particulateConc
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
            Thread.Sleep(1000)
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
                    '杭州春来科技有限公司 DMS-100  型颗粒物测量仪管控
                    a = master.ReadHoldingRegisters(&H1, &H1C, 2)
                    _particulateConc = ModbusUtility.GetSingle(a(0), a(1))
                    ' _particulateConc = 1.12
                    ' WriteLogToText(Now.ToString("yyyyMMdd"), "Particulate", _particulateConc)
                Catch ex As Exception
                    receivedatavalue1 = Now.ToString("ss")
                    '  _particulateConc = 1.23
                    Return False
                End Try
            Case Else
                sendCount = 1
        End Select
        Return True
    End Function
End Class
