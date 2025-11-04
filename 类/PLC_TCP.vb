Imports Modbus.Device
Imports Modbus.Utility
Imports System.IO
Imports System.Net.Sockets
Imports System.Threading

Public Class PLC_TCP

    Private master As ModbusIpMaster
    Private _reTryConnectTick As Integer

    Private WithEvents tcpclient As TcpClient
    Private tcpsenddata As Thread
    Private ReadOnly CheckOnLine As Thread

    Public setholdflag As Boolean
    Public holdregadd As UShort
    Public holdregdata As UShort()

    Private _tempConc As Single = 0
    Private _pressConc As Single = 0
    Private _speedConc As Single = 0
    Private _humidConc As Single = 0

    '系统工作状态，包括系统维护，系统运行，皮托管反吹，烟气反吹
    Structure status
        Dim pitotFlow As Boolean
        Dim gasFlow As Boolean
        Dim systemMain As Boolean
        Dim systemRun As Boolean
        Dim systemZero As Boolean
        Dim systemCal As Boolean
    End Structure

    Public PLCStatus As status

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

    Public ReadOnly Property HUMIDConc() As Single
        Get
            Return _humidConc
        End Get

    End Property

    Private Sub ConnectToServer()

        Try
            'Dim sToken() As String
            'sToken = _ipAndport.Split(",")
            If tcpclient IsNot Nothing Then tcpclient.Close()
            tcpclient = New TcpClient("192.168.8.110", 502)

            master = ModbusIpMaster.CreateIp(tcpclient)
            With master
                .Transport.ReadTimeout = 200
                .Transport.WriteTimeout = 200
                .Transport.WaitToRetryMilliseconds = 200
                .Transport.Retries = 2
            End With
            'tcpclient.Connect("127.0.0.1", 5000)

            'TcpReceiveData = New Thread(AddressOf ReceiveData)
            'With TcpReceiveData
            '    .IsBackground = True
            '    .Start()
            'End With

        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' 断开与服务器的连接
    ''' </summary>
    Public Sub CloseToServer()
        Try

            If tcpclient IsNot Nothing Then
                If tcpclient.Connected Then tcpclient.Close()
            End If
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' 网络断线重连
    ''' </summary>
    Private Sub RetryConnect()

        While True
            '当网络中断时，等待10秒，
            If (tcpclient Is Nothing) OrElse (tcpclient.Client Is Nothing) OrElse (tcpclient.Connected = False) Then
                _reTryConnectTick += 1
            Else
                _reTryConnectTick = 0
            End If
            '然后进行连接
            If _reTryConnectTick > 10 Then
                _reTryConnectTick = 0
                ConnectToServer()
            End If

            Thread.Sleep(1000)
        End While

    End Sub

    ''' <summary>
    ''' 是否在线TCP
    ''' </summary>
    ''' <param name="tcpclient"></param>
    ''' <returns></returns>
    Private Function IsConnected(tcpclient As TcpClient) As Boolean

        Try

            If tcpclient.Client.Connected AndAlso (tcpclient.Client.Poll(0, SelectMode.SelectRead)) Then
                If tcpclient.Client.Available > 0 Then
                    Return True
                Else
                    tcpclient.Close()
                    Return False
                End If
            End If

        Catch ex As Exception
            Return False
        End Try
        Return tcpclient.Client.Connected

    End Function

    ''' <summary>
    ''' 网络重连线程
    ''' </summary>
    Sub New(ipAndPort As String)
        '_ipAndport = ipAndPort

        'Call ConnectToServer()

        'master = ModbusIpMaster.CreateIp(tcpclient)
        'With master
        '    .Transport.ReadTimeout = 200
        '    .Transport.WriteTimeout = 200
        '    .Transport.WaitToRetryMilliseconds = 200
        '    .Transport.Retries = 2
        'End With

        CheckOnLine = New Thread(AddressOf RetryConnect)
        With CheckOnLine
            .IsBackground = True
            .Start()
        End With

        tcpsenddata = New Thread(AddressOf SendData)
        With tcpsenddata
            .IsBackground = True
            .Start()
        End With
    End Sub

    Private Sub analyStatus(status As Integer)

        If (status And (2 ^ 0)) <> 0 Then
            PLCStatus.pitotFlow = True
        Else
            PLCStatus.pitotFlow = False
        End If

        If (status And (2 ^ 1)) <> 0 Then
            PLCStatus.gasFlow = True
        Else
            PLCStatus.gasFlow = False
        End If

        If (status And (2 ^ 2)) <> 0 Then
            PLCStatus.systemRun = True
        Else
            PLCStatus.systemRun = False
        End If

        If (status And (2 ^ 3)) <> 0 Then
            PLCStatus.systemMain = True
        Else
            PLCStatus.systemMain = False
        End If
    End Sub

    Private Sub SendData()
        'If tcpclient Is Nothing OrElse tcpclient.Client Is Nothing OrElse tcpclient.Client.Connected = False Then Exit Sub
        'Dim ns As NetworkStream = tcpclient.GetStream()
        'Dim ws As New BinaryWriter(ns)
        Static count As Integer
        Dim a As UShort()

        Dim b(3) As Single
        While True
            Try
                If master IsNot Nothing Then
                    If Not setholdflag Then
                        Select Case count
                            Case 0
                                a = master.ReadHoldingRegisters(1, 0, 8)
                                _tempConc = ModbusUtility.GetSingle(a(0), a(1))
                                _pressConc = ModbusUtility.GetSingle(a(2), a(3))
                                _speedConc = ModbusUtility.GetSingle(a(4), a(5))
                                _humidConc = ModbusUtility.GetSingle(a(6), a(7))
                            Case 1
                                a = master.ReadHoldingRegisters(1, 105, 1)
                                analyStatus(a(0))
                            Case 2

                            Case Else
                                count = -1
                        End Select
                    Else
                        master.WriteMultipleRegisters(1, holdregadd, holdregdata)
                        setholdflag = False
                    End If
                End If

            Catch ex As Exception

            End Try
            count += 1

            Thread.Sleep(500)
        End While
        'While True
        '    Try
        '        '定时发送心跳包，默认5分钟
        '        Dim buff As Byte()

        '        buff = System.Text.Encoding.GetEncoding("gb2312").GetBytes("我是柏盛")
        '        ws.Write(buff)
        '        'Call SendHeChaData_2062()
        '    Catch ex As Exception

        '    End Try

        '    Thread.Sleep(1000)
        'End While

    End Sub
End Class
