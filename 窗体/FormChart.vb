Public Class FormChart

    Private selectParameter As String

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Try
            Chart1.Legends(0).Enabled = CheckBox2.Checked
        Catch ex As Exception

        End Try


    End Sub
    Private Sub SetButtonEnableStatus(isQueryEable As Boolean)
        Button1.Enabled = isQueryEable '浏览
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sqlHelper As SQLiteHelper
        Dim strStartDate As String
        Dim strEndDate As String
        Dim listx As New List(Of String)
        Dim list1 As New List(Of Single)

        SetButtonEnableStatus(False)
        selectNum()

        strStartDate = DateTimePicker1.Value.ToString("yyyy-MM-dd") & " " & DateTimePicker2.Value.ToString("HH:mm") & ":00"
        strEndDate = DateTimePicker3.Value.ToString("yyyy-MM-dd") & " " & DateTimePicker4.Value.ToString("HH:mm") & ":00"
        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            SetButtonEnableStatus(True)
            Exit Sub
        End If

        Dim sql As String = "SELECT strftime('%m-%d %H:%m',samplingtime) as 时间 " & selectParameter & ",remark FROM 分钟数据 where samplingtime>='" & strStartDate & "' and samplingtime <='" & strEndDate & "' "

        sqlHelper.SelectToDt(sql)

        If Not sqlHelper.SQLStatus Then
            SetButtonEnableStatus(True)
            sqlHelper.Close()
            Exit Sub
        End If

        Chart1.DataSource = sqlHelper.DatasDt
        Chart1.Series(0).XValueMember = "时间"
        Chart1.Series(0).YValueMembers = "甲烷"

        '甲烷排量的XValueMember
        Chart1.Series(1).XValueMember = "时间"
        Chart1.Series(1).YValueMembers = "甲烷排量"

        Chart1.Series(2).XValueMember = "时间"
        Chart1.Series(2).YValueMembers = "非甲烷"

        Chart1.Series(3).XValueMember = "时间"
        Chart1.Series(3).YValueMembers = "非甲烷排量"

        Chart1.Series(4).XValueMember = "时间"
        Chart1.Series(4).YValueMembers = "总烃"

        Chart1.Series(5).XValueMember = "时间"
        Chart1.Series(5).YValueMembers = "总烃排量"

        Chart1.Series(6).XValueMember = "时间"
        Chart1.Series(6).YValueMembers = "氧气"

        Chart1.Series(7).XValueMember = "时间"
        Chart1.Series(7).YValueMembers = "温度"

        Chart1.Series(8).XValueMember = "时间"
        Chart1.Series(8).YValueMembers = "压力"

        Chart1.Series(9).XValueMember = "时间"
        Chart1.Series(9).YValueMembers = "流速"

        Chart1.Series(10).XValueMember = "时间"
        Chart1.Series(10).YValueMembers = "标况流量"

        Chart1.Series(11).XValueMember = "时间"
        Chart1.Series(11).YValueMembers = "湿度"

        Chart1.ChartAreas(0).CursorX.IsUserSelectionEnabled = True
        Chart1.DataBind()

        ToolStripStatusLabel1.Text = "查询到数据：" & sqlHelper.DatasDt.Rows.Count & "  条"
        sqlHelper.Close()
        SetButtonEnableStatus(True)

    End Sub


    Private Sub selectNum()

        Dim strbuf As String = ""
        strbuf &= ",ch4conc as 甲烷"
        strbuf &= ",ch4conc_p as  甲烷排量"
        strbuf &= ",nmhcconc as 非甲烷"
        strbuf &= ",nmhcconc_p as 非甲烷排量"
        strbuf &= ",tchconc as 总烃"
        strbuf &= ",tchconc_p as 总烃排量"
        strbuf &= ",o2conc as 氧气"
        strbuf &= ",tempconc as 温度"
        strbuf &= ",pressconc as 压力"
        strbuf &= ",speedconc as 流速"
        strbuf &= ",flowconc_b as 标况流量"
        strbuf &= ",humiconc as 湿度"
        selectParameter = strbuf
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


    Private Sub CheckedListBox1_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles CheckedListBox1.ItemCheck

        Chart1.Series(e.Index).Enabled = e.NewValue

    End Sub

    Private Sub FormChart_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = Now.Date.ToString("yyyy-MM-dd")
        DateTimePicker3.Value = Now.Date.ToString("yyyy-MM-dd")

        DateTimePicker2.Value = Now.Date.ToString("yyyy-MM-dd") & " 00:00:00"
        DateTimePicker4.Value = Now.Date.ToString("yyyy-MM-dd") & " 23:59:59"

        For i As Integer = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemChecked(i, True)
        Next
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim savefile As New SaveFileDialog

        With savefile
            .FileName = Now.ToString("yyyyMMddHHmm")
            .Filter = "jpeg files(*.jpg)|*.jpeg|bmp files(*.bmp)|*.bmp|png files(*.png)|*.png"
            .RestoreDirectory = False
        End With

        If savefile.ShowDialog() = DialogResult.Cancel Then
            Exit Sub
        End If

        Select Case savefile.FilterIndex
            Case 1
                Chart1.SaveImage(savefile.FileName, DataVisualization.Charting.ChartImageFormat.Jpeg)
            Case 2
                Chart1.SaveImage(savefile.FileName, DataVisualization.Charting.ChartImageFormat.Bmp)
            Case 3
                Chart1.SaveImage(savefile.FileName, DataVisualization.Charting.ChartImageFormat.Png)
        End Select

    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged

    End Sub

End Class