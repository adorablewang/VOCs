Imports System.IO

Public Class FormMinuteReport
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        SetButtonEnableStatus(False, False)
        Dim tr As Threading.Thread
        tr = New Threading.Thread(AddressOf dd)
        tr.IsBackground = True
        tr.Start()

    End Sub

    Private Sub dd()
        Dim dl As New Action(AddressOf viewminutereport)
        Me.Invoke(dl)
    End Sub
    Private Sub SetButtonEnableStatus(isQueryEable As Boolean, isExportEable As Boolean)
        Button1.Enabled = isQueryEable '浏览
        Button2.Enabled = isExportEable '导出      
    End Sub
    Private Sub viewminutereport()
        Dim sqlHelper As SQLiteHelper
        Dim strStartDate As String
        Dim strEndDate As String

        strStartDate = DateTimePicker1.Value.ToString("yyyy-MM-dd")
        strEndDate = DateTimePicker2.Value.ToString("yyyy-MM-dd")
        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            SetButtonEnableStatus(True, True)
            Exit Sub
        End If


        Dim sql As String = "SELECT samplingtime as 时间,ch4conc_h as 甲烷（湿基）,ch4conc as 甲烷（干基）,ch4conc_z as 甲烷（折算）,ch4conc_p as 甲烷（排放量）,
                            nmhcconc_h as 非甲烷（湿基）,nmhcconc as 非甲烷（干基）,nmhcconc_z as 非甲烷（折算）,nmhcconc_p as 非甲烷（排放量）,
                            tchconc_h as 总烃（湿基）,tchconc as 总烃（干基）,tchconc_z as 总烃（折算）,tchconc_p as 总烃（排放量）,
                            o2conc_h as 氧气（湿基）,o2conc as 氧气（干基）,tempconc as 烟温,pressconc as 静压,speedconc as 流速,
                            flowconc_g as 流量（工况）,flowconc_b as 流量（标况）,humiconc as 湿度 FROM 分钟数据 where samplingtime>='" & strStartDate & "' and samplingtime <='" & strEndDate & "' "

        sqlHelper.SelectToDt(sql)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            SetButtonEnableStatus(True, True)
            Exit Sub
        End If

        'For i = 0 To sqlHelper.DatasDs.Tables.Count - 1

        With DataGridView1
            .DataSource = Nothing
            .DataSource = sqlHelper.DatasDt
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End With

        InitializeDataGridView()

        'ToolStripStatusLabel1.Text = "查询到数据：" & sqlHelper.DatasDt.Rows.Count & "  条"
        sqlHelper.Close()
        SetButtonEnableStatus(True, True)
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
        SetButtonEnableStatus(False, False)

        Dim st1 As Stream = Nothing
        Dim savefile As New SaveFileDialog

        savefile.Filter = "CSV(*.csv)|*.csv|文本文件(*.txt)|*.txt"
        savefile.InitialDirectory = Application.StartupPath
        savefile.RestoreDirectory = True
        savefile.FileName = Now.ToString("yyyyMMddHHmm")

        If savefile.ShowDialog() = DialogResult.Cancel Then

            SetButtonEnableStatus(True, True)
            Exit Sub
        End If

        Dim f As String = savefile.FileName

        st1 = savefile.OpenFile
        Export_csv(DataGridView1, st1)
        st1.Dispose()
        SetButtonEnableStatus(True, True)
    End Sub
End Class