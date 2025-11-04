Imports System.IO


Public Class FormReport

    Private selectParameter As String


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sqlHelper As SQLiteHelper
        Dim strStartDate As String
        Dim strEndDate As String
        Dim readSqlRecordCnt As Integer
        '禁用浏览
        SetButtonEnableStatus(False, False, False)
        selectNum()
        If selectParameter = "" Then
            MessageBox.Show("当前左侧Item未有选中项,请至少勾选一项后重试!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            SetButtonEnableStatus(True, True, True)
            Exit Sub
        End If


        strStartDate = DateTimePicker1.Value.ToString("yyyy-MM-dd") & " " & DateTimePicker2.Value.ToString("HH:mm") & ":00"
        strEndDate = DateTimePicker3.Value.ToString("yyyy-MM-dd") & " " & DateTimePicker4.Value.ToString("HH:mm") & ":00"
        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            MessageBox.Show("当前数据库状态异常!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            '禁用浏览
            SetButtonEnableStatus(True, True, True)
            Exit Sub
        End If

        '读取当前记录查询总数
        Dim sqlCount As String = "SELECT  count(*) FROM 分钟数据 where samplingtime>='" & strStartDate & "' and samplingtime <='" & strEndDate & "' "
        sqlHelper.SelectToDt(sqlCount)
        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            MessageBox.Show("当前数据库查询失败!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            SetButtonEnableStatus(True, True, True)
            Exit Sub
        End If
        readSqlRecordCnt = sqlHelper.DatasDt.Rows(0).Item(0)
        Console.WriteLine($" {readSqlRecordCnt}")

        If readSqlRecordCnt > 3000 Then
            MessageBox.Show("当前查询记录过大,请缩减查询时间间隔后重试!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            SetButtonEnableStatus(True, True, True)
            Exit Sub
        End If
        Dim sql As String = "SELECT samplingtime as 时间 " & selectParameter & ",remark as 备注 FROM 分钟数据 where samplingtime>='" & strStartDate & "' and samplingtime <='" & strEndDate & "'  ORDER BY samplingtime DESC "
        sqlHelper.SelectToDt(sql)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            MessageBox.Show("当前数据库查询失败!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            SetButtonEnableStatus(True, True, True)
            Exit Sub
        End If

        'For i = 0 To sqlHelper.DatasDs.Tables.Count - 1

        With DataGridView1
            .DataSource = Nothing
            .DataSource = sqlHelper.DatasDt
            '.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End With

        ToolStripStatusLabel1.Text = "查询到数据：" & sqlHelper.DatasDt.Rows.Count & "  条"
        sqlHelper.Close()

        InitializeDataGridView()

        DisplayColumnHead()
        'Dim strCap As String()
        'ReDim strCap(1)

        'strCap(1) = "时间,甲烷,甲烷,甲烷,非甲烷总烃,非甲烷总烃,非甲烷总烃,总烃,总烃,总烃," &
        '   "流量" & vbCrLf & "m3/3,O2" & vbCrLf & "%,温度" & vbCrLf & "℃,湿度" & vbCrLf & "%,负荷,备注"                        '这是最下面一行列标题的内容，必须注意，个数一定和实际列对应
        'strCap(0) = "时间,mg/m3,折算" & vbCrLf & "mg/m3,kg/h,mg/m3,折算" & vbCrLf & "mg/m3,kg/h,mg/m3,折算" & vbCrLf & "mg/m3,kg/h," &
        '            "流量" & vbCrLf & "m3/3,O2" & vbCrLf & "%,温度" & vbCrLf & "℃,湿度" & vbCrLf & "%,负荷,备注"           '第2行的内容
        ''strCap(2) = "HCol1,HCol2,HCol2,HCol2,HCol5,HCol6,HCol6,HCol6,HCol6,Col10"       '第3行的内容
        'DgV_MulCapRowMerge1.ColumnSCaption = strCap
        'DgV_MulCapRowMerge1.ColumnDeep = 2
        SetButtonEnableStatus(True, True, True)
    End Sub

    Private Sub InitializeDataGridView()

        With DataGridView1
            '.Columns(0).Visible = False
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            .AllowUserToAddRows = False
            .AllowUserToResizeRows = False
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            .ColumnHeadersDefaultCellStyle.Font = New Font("宋体", 14)
            .DefaultCellStyle.Font = New Font("宋体", 14)
            .RowsDefaultCellStyle.SelectionForeColor = SystemColors.HighlightText
            .RowsDefaultCellStyle.SelectionBackColor = SystemColors.Highlight
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            '设置奇数行的颜色
            .AlternatingRowsDefaultCellStyle.BackColor = Color.PaleGreen
        End With

        'For i As Integer = 0 To DataGridView1.Columns.Count - 1
        '    With DataGridView1
        '        .Columns(i).DefaultCellStyle.Font = New Font("宋体", 14)
        '    End With
        'Next

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim st1 As Stream = Nothing
        Dim savefile As New SaveFileDialog
        SetButtonEnableStatus(False, False, False)
        savefile.Filter = "CSV(*.csv)|*.csv|文本文件(*.txt)|*.txt"
        savefile.InitialDirectory = Application.StartupPath
        savefile.RestoreDirectory = True
        savefile.FileName = Now.ToString("yyyyMMddHHmm")

        If savefile.ShowDialog() = DialogResult.Cancel Then
            SetButtonEnableStatus(True, True, True)
            Exit Sub
        End If

        Dim f As String = savefile.FileName

        st1 = savefile.OpenFile
        Export_csv(DataGridView1, st1)
        st1.Dispose()
        SetButtonEnableStatus(True, True, True)
    End Sub
    Private Sub SetButtonEnableStatus(isQueryEable As Boolean, isExportEable As Boolean, isExitEable As Boolean)
        Button1.Enabled = isQueryEable '浏览
        Button2.Enabled = isExportEable '导出
        Button3.Enabled = isExitEable '退出

    End Sub
    Private Sub selectNum()
        Dim i As Integer
        Dim strbuf As String = ""
        '新增颗粒物三项指标, add by wx at 2024.08.22 10:03
        For i = 0 To CheckedListBox1.Items.Count - 1
            If CheckedListBox1.GetItemChecked(i) Then
                If i = 0 Then strbuf = strbuf & ",ch4conc as 甲烷"
                If i = 1 Then strbuf = strbuf & ",ch4conc_z as 甲烷折算"
                If i = 2 Then strbuf = strbuf & ",ch4conc_p as  甲烷排量"
                If i = 3 Then strbuf = strbuf & ",nmhcconc as 非甲烷"
                If i = 4 Then strbuf = strbuf & ",nmhcconc_z as 非甲烷折算"
                If i = 5 Then strbuf = strbuf & ",nmhcconc_p as 非甲烷排量"
                If i = 6 Then strbuf = strbuf & ",tchconc as 总烃"
                If i = 7 Then strbuf = strbuf & ",tchconc_z as 总烃折算"
                If i = 8 Then strbuf = strbuf & ",tchconc_p as 总烃排量"
                If i = 9 Then strbuf = strbuf & ",particulateconc as 颗粒物"
                If i = 10 Then strbuf = strbuf & ",particulateconc_z as 颗粒物折算"
                If i = 11 Then strbuf = strbuf & ",particulateconc_p as 颗粒物排量"
                If i = 12 Then strbuf = strbuf & ",o2conc as 氧气"
                If i = 13 Then strbuf = strbuf & ",tempconc as 温度"
                If i = 14 Then strbuf = strbuf & ",pressconc as 压力"
                If i = 15 Then strbuf = strbuf & ",speedconc as 流速"
                If i = 16 Then strbuf = strbuf & ",flowconc_g as 工况流量"
                If i = 17 Then strbuf = strbuf & ",flowconc_b as 标况流量"
                If i = 18 Then strbuf = strbuf & ",humiconc as 湿度"
                If i = 19 Then strbuf = strbuf & ",fuheconc as 大气压"
            End If
        Next

        selectParameter = strbuf
    End Sub

    Private Sub DisplayColumnHead()


        Dim i As Integer
        Dim strbuf As String = ""
        Dim tokens As String()

        For i = 0 To CheckedListBox1.Items.Count - 1
            If CheckedListBox1.GetItemChecked(i) Then
                If i = 0 Then strbuf = strbuf & "甲烷" & vbCrLf & "mg/m³"
                If i = 1 Then strbuf = strbuf & ",甲烷折算" & vbCrLf & "mg/m³"
                If i = 2 Then strbuf = strbuf & ",甲烷排量" & vbCrLf & "m³/h"
                If i = 3 Then strbuf = strbuf & ",非甲烷" & vbCrLf & "mg/m³"
                If i = 4 Then strbuf = strbuf & ",非甲烷折算" & vbCrLf & "mg/m³"
                If i = 5 Then strbuf = strbuf & ",非甲烷排量" & vbCrLf & "m³/h"
                If i = 6 Then strbuf = strbuf & ",总烃" & vbCrLf & "mg/m³"
                If i = 7 Then strbuf = strbuf & ",总烃折算" & vbCrLf & "mg/m³"
                If i = 8 Then strbuf = strbuf & ",总烃排量" & vbCrLf & "m³/h"
                If i = 9 Then strbuf = strbuf & ",颗粒物" & vbCrLf & "mg/m³"
                If i = 10 Then strbuf = strbuf & ",颗粒物折算" & vbCrLf & "mg/m³"
                If i = 11 Then strbuf = strbuf & ",颗粒物排量" & vbCrLf & "m³/h"
                If i = 12 Then strbuf = strbuf & ",氧气" & vbCrLf & "%"
                If i = 13 Then strbuf = strbuf & ",温度" & vbCrLf & "℃"
                If i = 14 Then strbuf = strbuf & ",压力" & vbCrLf & "kpa"
                If i = 15 Then strbuf = strbuf & ",流速" & vbCrLf & "m/s"
                If i = 16 Then strbuf = strbuf & ",工况流量" & vbCrLf & "m³/h"
                If i = 17 Then strbuf = strbuf & ",标况流量" & vbCrLf & "m³/h"
                If i = 18 Then strbuf = strbuf & ",湿度" & vbCrLf & "%"
                If i = 19 Then strbuf = strbuf & ",大气压"
            End If
        Next

        selectParameter = strbuf
        If strbuf.Substring(0, 1) = "," Then
            strbuf = strbuf.Substring(1)
        End If
        tokens = strbuf.Split(",")

        For i = 1 To tokens.Length - 1
            DataGridView1.Columns(i).HeaderText = tokens(i - 1)
        Next

        'DgV_MulCapRowMerge1.Columns(0).HeaderText = "日期"
        'DgV_MulCapRowMerge1.Columns(1).HeaderText = "甲烷" & vbCrLf & "t/m"
        'DgV_MulCapRowMerge1.Columns(2).HeaderText = "非甲烷总烃" & vbCrLf & "t/m"
        'DgV_MulCapRowMerge1.Columns(3).HeaderText = "总烃" & vbCrLf & "t/m"
        'DgV_MulCapRowMerge1.Columns(4).HeaderText = "流量" & vbCrLf & "×10" & "⁴" & vbCrLf & "m³/h,"
        'DgV_MulCapRowMerge1.Columns(5).HeaderText = "O₂" & vbCrLf & "%"
        'DgV_MulCapRowMerge1.Columns(6).HeaderText = "温度" & vbCrLf & "℃"
        'DgV_MulCapRowMerge1.Columns(7).HeaderText = "湿度" & vbCrLf & "%"
        'DgV_MulCapRowMerge1.Columns(8).HeaderText = "大气压"
        'DgV_MulCapRowMerge1.Columns(9).HeaderText = "备注"
    End Sub


    Private Sub FormReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        DateTimePicker1.Value = Now.Date.ToString("yyyy-MM-dd")
        DateTimePicker3.Value = Now.Date.ToString("yyyy-MM-dd")

        DateTimePicker2.Value = Now.Date.ToString("yyyy-MM-dd") & " 00:00:00"
        DateTimePicker4.Value = Now.Date.ToString("yyyy-MM-dd") & " 23:59:59"
        SetButtonEnableStatus(True, True, True) '设置该节界面按钮使能状态
        For i As Integer = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemChecked(i, True)
        Next

        ToolStripStatusLabel2.Text = "A：调试，K：核查比对，N：采样，B：反吹，C：校准，M：维护，D：故障，T：超标，H：有效数据不足"
        ToolStripStatusLabel1.Text = "查询到数据：  条"
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Dim i As Integer
        If CheckBox1.Checked Then
            For i = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(i, True)
            Next
        Else
            For i = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(i, False)
            Next
        End If
    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub
End Class