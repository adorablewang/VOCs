Imports Ini.Net

Public Class FormUartSetting
    Private Sub FormUartSetting_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim strbuf As String
        Select Case ListBox1.SelectedIndex
            Case 0
                strbuf = analyzerdata.SerialConfig
            Case 1
                strbuf = plcdata.SerialConfig
            Case 2
                strbuf = o2data.SerialConfig
            Case 3
                strbuf = datasamplinganalyzer.PortName & "," & datasamplinganalyzer.BaudRate & ",N," & datasamplinganalyzer.DataBits & "," & datasamplinganalyzer.StopBits
            Case 4
                strbuf = particulatedata.SerialConfig '"COM11,9600,N,8,1"

            Case Else
                Exit Sub
        End Select

        Dim tokens As String()
        tokens = strbuf.Split(",")
        ComboBox1.Text = tokens(0)
        ComboBox2.Text = tokens(1)
        ComboBox4.Text = tokens(3)
        ComboBox5.Text = tokens(4)

        If tokens(2).ToUpper = "N" Then ComboBox3.SelectedIndex = 0

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim InitializeName As New IniFile(IniPathName)
        Dim strbuf As String
        Dim s As String = ""

        Select Case ComboBox3.SelectedIndex
            Case 0
                s = "N"
            Case 1
                s = "O"
            Case 2
                s = "E"
            Case 3
                s = "M"
            Case 4
                s = "S"
        End Select

        strbuf = ComboBox1.Text & "," & ComboBox2.Text & "," & s & "," & ComboBox4.Text & "," & ComboBox5.Text

        Select Case ListBox1.SelectedIndex
            Case 0
                InitializeName.WriteString("通讯设置", "甲烷", strbuf)
            Case 1
                InitializeName.WriteString("通讯设置", "PLC", strbuf)
            Case 2
                InitializeName.WriteString("通讯设置", "氧气", strbuf)
            Case 3
                InitializeName.WriteString("通讯设置", "数采仪", strbuf)
            Case 4
                InitializeName.WriteString("通讯设置", "颗粒物", strbuf)
        End Select


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

End Class