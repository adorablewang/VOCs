Public Class FormChangeLog
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Select Case TabControl1.SelectedIndex
            Case 0
                DisplayKeyParameterLog()
            Case 1
                DisplayRangeParameterLog()
            Case 2
                '根据ComboBox2的选择项进行显示
                If ComboBox2.SelectedIndex = 0 Then
                    DisplaySystemStart()
                ElseIf ComboBox2.SelectedIndex = 1 Then
                    displayLoginRecord()
                ElseIf ComboBox2.SelectedIndex = 2 Then
                    displayCalibrationRecord()
                End If
            Case 3
                DiplayAlarmLog()
        End Select

    End Sub

    Private Sub FormChangeLog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = Now.Date.ToString("yyyy-MM-dd")
        DateTimePicker3.Value = Now.Date.ToString("yyyy-MM-dd")

        DateTimePicker2.Value = Now.Date.ToString("yyyy-MM-dd") & " 00:00:00"
        DateTimePicker4.Value = Now.Date.ToString("yyyy-MM-dd") & " 23:59:59"

        ComboBox2.SelectedIndex = 0
    End Sub

    ' 定义事件处理程序
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ' 这里编写当选择发生变化时想要执行的代码
        Dim comboBox As ComboBox = CType(sender, ComboBox)
        Dim newIndex As Integer = comboBox.SelectedIndex
    End Sub

    ''' <summary>
    ''' 显示登录记录
    ''' </summary>
    Private Sub displayLoginRecord()
        Dim sqlHelper As SQLiteHelper
        Dim strStartDate As String
        Dim strEndDate As String

        strStartDate = DateTimePicker1.Value.ToString("yyyy-MM-dd") & " " & DateTimePicker2.Value.ToString("HH:mm") & ":00"
        strEndDate = DateTimePicker3.Value.ToString("yyyy-MM-dd") & " " & DateTimePicker4.Value.ToString("HH:mm") & ":00"
        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        Dim sql As String = "SELECT samplingtime as 时间,username as 用户,info as 记录,infotype FROM 事件记录 
                            where samplingtime>='" & strStartDate & "' and samplingtime <='" & strEndDate & "' and infotype = 1"

        sqlHelper.SelectToDt(sql)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If

        'For i = 0 To sqlHelper.DatasDs.Tables.Count - 1

        With DataGridView3

            .DataSource = Nothing
            .DataSource = sqlHelper.DatasDt
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            .DefaultCellStyle.Font = New Font("宋体", 14)
            .AlternatingRowsDefaultCellStyle.BackColor = Color.PaleGreen
            .Columns(3).Visible = False
        End With

        sqlHelper.Close()

    End Sub

    ''' <summary>
    ''' 显示校准记录
    ''' </summary>
    Private Sub displayCalibrationRecord()
        Dim sqlHelper As SQLiteHelper
        Dim strStartDate As String
        Dim strEndDate As String

        strStartDate = DateTimePicker1.Value.ToString("yyyy-MM-dd") & " " & DateTimePicker2.Value.ToString("HH:mm") & ":00"
        strEndDate = DateTimePicker3.Value.ToString("yyyy-MM-dd") & " " & DateTimePicker4.Value.ToString("HH:mm") & ":00"
        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        Dim sql As String = "SELECT samplingtime as 时间,username as 用户,info as 记录,infotype FROM 事件记录 
                            where samplingtime>='" & strStartDate & "' and samplingtime <='" & strEndDate & "' and infotype = 2"

        sqlHelper.SelectToDt(sql)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If

        'For i = 0 To sqlHelper.DatasDs.Tables.Count - 1

        With DataGridView3

            .DataSource = Nothing
            .DataSource = sqlHelper.DatasDt
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            .DefaultCellStyle.Font = New Font("宋体", 14)
            .AlternatingRowsDefaultCellStyle.BackColor = Color.PaleGreen
            .Columns(3).Visible = False
        End With

        sqlHelper.Close()

    End Sub

    ''' <summary>
    ''' 显示所有记录
    ''' </summary>
    Private Sub DisplaySystemStart()
        Dim sqlHelper As SQLiteHelper
        Dim strStartDate As String
        Dim strEndDate As String

        strStartDate = DateTimePicker1.Value.ToString("yyyy-MM-dd") & " " & DateTimePicker2.Value.ToString("HH:mm") & ":00"
        strEndDate = DateTimePicker3.Value.ToString("yyyy-MM-dd") & " " & DateTimePicker4.Value.ToString("HH:mm") & ":00"
        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        Dim sql As String = "SELECT samplingtime as 时间,username as 用户,info as 记录 FROM 事件记录 
                            where samplingtime>='" & strStartDate & "' and samplingtime <='" & strEndDate & "' "

        sqlHelper.SelectToDt(sql)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If

        'For i = 0 To sqlHelper.DatasDs.Tables.Count - 1

        With DataGridView3

            .DataSource = Nothing
            .DataSource = sqlHelper.DatasDt
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            .DefaultCellStyle.Font = New Font("宋体", 14)
            .AlternatingRowsDefaultCellStyle.BackColor = Color.PaleGreen
        End With

        sqlHelper.Close()
    End Sub

    Private Sub DiplayAlarmLog()
        Dim sqlHelper As SQLiteHelper
        Dim strStartDate As String
        Dim strEndDate As String

        strStartDate = DateTimePicker1.Value.ToString("yyyy-MM-dd") & " " & DateTimePicker2.Value.ToString("HH:mm") & ":00"
        strEndDate = DateTimePicker3.Value.ToString("yyyy-MM-dd") & " " & DateTimePicker4.Value.ToString("HH:mm") & ":00"
        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        Dim sql As String = "SELECT samplingtime,ch4conc,ch4conc_z,ch4conc_p,
                            nmhcconc,nmhcconc_z,nmhcconc_p,tchconc,tchconc_z,tchconc_p,
                            flowconc_b,o2conc,tempconc,humiconc,fuheconc,remark FROM 小时数据 where samplingtime>='" & strStartDate & "' and samplingtime <='" & strEndDate & "' "

        Select Case ComboBox1.SelectedIndex
            Case 0
                sql = "select samplingtime as 时间,ch4conc_z as 超标浓度 from 小时数据 where samplingtime>='" & strStartDate & "' and samplingtime <='" & strEndDate & "' and ch4conc_z>" & alarmParameter.ch4alarmup & " "
            Case 1
                sql = "select samplingtime as 时间,nmhcconc_z as 超标浓度 from 小时数据 where samplingtime>='" & strStartDate & "' and samplingtime <='" & strEndDate & "' and nmhcconc_z>" & alarmParameter.ch4alarmup & " "
            Case 2
                sql = "select samplingtime as 时间,tchconc_z as 超标浓度 from 小时数据 where samplingtime>='" & strStartDate & "' and samplingtime <='" & strEndDate & "' and tchconc_z>" & alarmParameter.ch4alarmup & " "
        End Select

        sqlHelper.SelectToDt(sql)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If

        'For i = 0 To sqlHelper.DatasDs.Tables.Count - 1
        Dim dt As DataTable
        Dim row As DataRow
        Dim str As DateTime
        dt = sqlHelper.DatasDt
        dt.Columns.Add("备注")
        For Each row In dt.Rows
            str = row(0)
            row(0) = str.AddHours(-1) & " --- " & row(0)
            row(2) = "超上限报警"
        Next

        With DataGridView4

            .DataSource = Nothing
            .DataSource = dt
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            .DefaultCellStyle.Font = New Font("宋体", 14)
            .AlternatingRowsDefaultCellStyle.BackColor = Color.PaleGreen
        End With

        sqlHelper.Close()
    End Sub

    Private Sub DisplayKeyParameterLog()
        Dim sqlHelper As SQLiteHelper
        Dim strStartDate As String
        Dim strEndDate As String

        strStartDate = DateTimePicker1.Value.ToString("yyyy-MM-dd") & " " & DateTimePicker2.Value.ToString("HH:mm") & ":00"
        strEndDate = DateTimePicker3.Value.ToString("yyyy-MM-dd") & " " & DateTimePicker4.Value.ToString("HH:mm") & ":00"
        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        Dim sql As String = "SELECT samplingtime as 时间,samplingarea as 烟道面积,press as 大气压,coeffo2 as 过剩系数,speed as 速度场系数,
                            pitol as 皮托管系数,limito2 as 折算限值,convertflag as 是否折算,
                            remark as 备注 FROM 关键参数记录 where samplingtime>='" & strStartDate & "' and samplingtime <='" & strEndDate & "' "

        sqlHelper.SelectToDt(sql)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If

        'For i = 0 To sqlHelper.DatasDs.Tables.Count - 1

        With DataGridView1

            .DataSource = Nothing
            .DataSource = sqlHelper.DatasDt
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            .DefaultCellStyle.Font = New Font("宋体", 14)
            .AlternatingRowsDefaultCellStyle.BackColor = Color.PaleGreen
        End With

        sqlHelper.Close()
    End Sub

    Private Sub DisplayRangeParameterLog()
        Dim sqlHelper As SQLiteHelper
        Dim strStartDate As String
        Dim strEndDate As String

        strStartDate = DateTimePicker1.Value.ToString("yyyy-MM-dd") & " " & DateTimePicker2.Value.ToString("HH:mm") & ":00"
        strEndDate = DateTimePicker3.Value.ToString("yyyy-MM-dd") & " " & DateTimePicker4.Value.ToString("HH:mm") & ":00"
        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        Dim sql As String = "SELECT samplingtime as 时间,tchlow as 总烃下限,tchup as 总烃上限,o2low as O2下限,o2up as O2上限,
                            templow as 温度下限,tempup as 温度上限,presslow as 压力下限,pressup as 压力上限,speedlow as 动压下限,speedup as 动压上限,
                            humilow as 湿度下限,humiup as 湿度上限 FROM 量程记录 where samplingtime>='" & strStartDate & "' and samplingtime <='" & strEndDate & "' "

        sqlHelper.SelectToDt(sql)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If

        'For i = 0 To sqlHelper.DatasDs.Tables.Count - 1

        With DataGridView2

            .DataSource = Nothing
            .DataSource = sqlHelper.DatasDt
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            .DefaultCellStyle.Font = New Font("宋体", 14)
            .AlternatingRowsDefaultCellStyle.BackColor = Color.PaleGreen
        End With

        sqlHelper.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub
End Class