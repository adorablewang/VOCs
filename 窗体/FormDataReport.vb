Imports Microsoft.Office.Interop

Public Class FormDataReport

    Private newstring As String = "m" & "³"

    Private newstring1 As String = "O" & "₂"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        SetButtonEnableStatus(False, False, False)
        DgV_MulCapRowMerge1.Columns.Clear()

        If RadioButton1.Checked Then
            ViewHourData()
        End If
        If RadioButton2.Checked Then
            ViewDailyData()
        End If

        If RadioButton3.Checked Then
            ViewMonthData()
        End If
        SetButtonEnableStatus(True, True, True)
    End Sub

    Private Sub InitializeDataGridView()

        With DgV_MulCapRowMerge1
            '.Columns(0).Visible = False
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

        'For i As Integer = 0 To DgV_MulCapRowMerge1.Columns.Count - 1
        '    With DgV_MulCapRowMerge1
        '        .Columns(i).DefaultCellStyle.Font = New Font("宋体", 14)
        '    End With
        'Next

    End Sub

    Private Sub ViewHourData()
        Dim sqlHelper As SQLiteHelper
        Dim strStartDate As String
        Dim strEndDate As String

        ' 2025.04.27 解决日报表查询没有"00:00"条目
        ' 原因过滤条件设置有误
        strStartDate = DateTimePicker1.Value.ToString("yyyy-MM-dd") & " 00:00"
        strEndDate = DateTimePicker1.Value.ToString("yyyy-MM-dd") & " 23:00"
        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        Dim sql As String = "SELECT samplingtime,ch4conc,ch4conc_z,ch4conc_p,
                            nmhcconc,nmhcconc_z,nmhcconc_p,tchconc,tchconc_z,tchconc_p,
                            flowconc_b,o2conc,tempconc,humiconc,fuheconc,remark FROM 小时数据 where samplingtime>='" & strStartDate & "' and samplingtime <='" & strEndDate & "'"

        sqlHelper.SelectToDt(sql)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If

        'For i = 0 To sqlHelper.DatasDs.Tables.Count - 1

        With DgV_MulCapRowMerge1
            .DataSource = Nothing
            .DataSource = sqlHelper.DatasDt
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End With

        sqlHelper.Close()

        InitializeDataGridView()

        Dim strCap As String()
        ReDim strCap(1)

        strCap(1) = "时间,甲烷,甲烷,甲烷,非甲烷总烃,非甲烷总烃,非甲烷总烃,总烃,总烃,总烃," &
           "流量" & vbCrLf & "m³" & "/h," & "O₂" & vbCrLf & "%,温度" & vbCrLf & "℃,湿度" & vbCrLf & "%,负荷,备注"                        '这是最下面一行列标题的内容，必须注意，个数一定和实际列对应
        strCap(0) = "时间,mg/m³" & ",折算" & vbCrLf & "mg/m³" & ",kg/h," & "mg/m³" & ",折算" & vbCrLf & "mg/m³" & ",kg/h," & "mg/m³" & ",折算" & vbCrLf & "mg/m³" & ",kg/h," &
                    "流量" & vbCrLf & "m³/h," & "O₂" & vbCrLf & "%,温度" & vbCrLf & "℃,湿度" & vbCrLf & "%,负荷,备注"           '第2行的内容
        'strCap(2) = "HCol1,HCol2,HCol2,HCol2,HCol5,HCol6,HCol6,HCol6,HCol6,Col10"       '第3行的内容
        DgV_MulCapRowMerge1.ColumnSCaption = strCap
        DgV_MulCapRowMerge1.ColumnDeep = 2


        For i As Integer = 1 To DgV_MulCapRowMerge1.Columns.Count - 1
            DgV_MulCapRowMerge1.Columns(i).Width = 80
        Next
    End Sub
    Private Sub SetButtonEnableStatus(isQueryEable As Boolean, isExportEable As Boolean, isExitEable As Boolean)
        Button1.Enabled = isQueryEable '浏览
        Button2.Enabled = isExportEable '导出
        Button3.Enabled = isExitEable '退出

    End Sub
    Private Sub ViewDailyData()
        Dim sqlHelper As SQLiteHelper
        Dim strStartDate As String
        Dim strEndDate As String

        strStartDate = DateTimePicker1.Value.ToString("yyyy-MM") & "-01 00:00:00 "
        strEndDate = DateTimePicker1.Value.AddMonths(1).ToString("yyyy-MM") & "-01 00:00:00 "
        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        Dim sql As String = "SELECT samplingtime,ch4conc,ch4conc_p,
                            nmhcconc,nmhcconc_p,tchconc,tchconc_p,
                            flowconc_b,o2conc,tempconc,humiconc,fuheconc,remark FROM 日数据 where samplingtime>='" & strStartDate & "' and samplingtime <'" & strEndDate & "' "

        sqlHelper.SelectToDt(sql)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If

        'For i = 0 To sqlHelper.DatasDs.Tables.Count - 1

        With DgV_MulCapRowMerge1
            .DataSource = Nothing
            .DataSource = sqlHelper.DatasDt
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End With

        sqlHelper.Close()

        InitializeDataGridView()

        Dim strCap As String()
        ReDim strCap(1)

        strCap(1) = "日期,甲烷,甲烷,非甲烷总烃,非甲烷总烃,总烃,总烃," &
           "流量" & vbCrLf & "×10⁴" & vbCrLf & "m³/h," & "O₂" & vbCrLf & "%,温度" & vbCrLf & "℃,湿度" & vbCrLf & "%,负荷,备注"                  '这是最下面一行列标题的内容，必须注意，个数一定和实际列对应
        strCap(0) = "日期,mg/m³" & ",t/d," & "mg/m³" & ",t/d," & "mg/m³" & ",kg/h," &
                    "流量" & vbCrLf & "×10⁴" & vbCrLf & "m³/h," & "O₂" & vbCrLf & "%,温度" & vbCrLf & "℃,湿度" & vbCrLf & "%,负荷,备注"         '第2行的内容
        'strCap(2) = "HCol1,HCol2,HCol2,HCol2,HCol5,HCol6,HCol6,HCol6,HCol6,Col10"                                                  '第3行的内容
        DgV_MulCapRowMerge1.ColumnSCaption = strCap
        DgV_MulCapRowMerge1.ColumnDeep = 2

        For i As Integer = 1 To DgV_MulCapRowMerge1.Columns.Count - 1
            DgV_MulCapRowMerge1.Columns(i).Width = 80
        Next
    End Sub

    Private Sub ViewMonthData()
        Dim sqlHelper As SQLiteHelper
        Dim strStartDate As String
        Dim strEndDate As String

        strStartDate = DateTimePicker1.Value.ToString("yyyy") & "-01-01 00:00:00 "
        strEndDate = DateTimePicker1.Value.ToString("yyyy") & "-12-31 23:59:59 "
        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        Dim sql As String = "SELECT samplingtime,ch4conc_p,
                            nmhcconc_p,tchconc_p,
                            flowconc_b,o2conc,tempconc,humiconc,fuheconc,remark FROM 月数据 
                            where samplingtime>='" & strStartDate & "' and samplingtime <='" & strEndDate & "'
                             order by samplingtime"

        sqlHelper.SelectToDt(sql)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If

        'For i = 0 To sqlHelper.DatasDs.Tables.Count - 1

        With DgV_MulCapRowMerge1
            .DataSource = Nothing
            .DataSource = sqlHelper.DatasDt
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End With

        sqlHelper.Close()

        InitializeDataGridView()

        Dim strCap As String()
        ReDim strCap(1)

        '注意这里数组strCap代表表格标题必须大小设置为2.
        'DGV_MulCapRowMerge控件的方法GetCaption中使用了“ReDim Preserve _ColCap(iRow, iCol)”
        '动态定义了2维数组。
        strCap(0) = ""
        strCap(1) = ""
        DgV_MulCapRowMerge1.ColumnSCaption = strCap
        DgV_MulCapRowMerge1.ColumnDeep = 1

        DgV_MulCapRowMerge1.Columns(0).HeaderText = "日期"
        DgV_MulCapRowMerge1.Columns(1).HeaderText = "甲烷" & vbCrLf & "t/m"
        DgV_MulCapRowMerge1.Columns(2).HeaderText = "非甲烷总烃" & vbCrLf & "t/m"
        DgV_MulCapRowMerge1.Columns(3).HeaderText = "总烃" & vbCrLf & "t/m"
        DgV_MulCapRowMerge1.Columns(4).HeaderText = "流量" & vbCrLf & "×10" & "⁴" & vbCrLf & "m³/h,"
        DgV_MulCapRowMerge1.Columns(5).HeaderText = "O₂" & vbCrLf & "%"
        DgV_MulCapRowMerge1.Columns(6).HeaderText = "温度" & vbCrLf & "℃"
        DgV_MulCapRowMerge1.Columns(7).HeaderText = "湿度" & vbCrLf & "%"
        DgV_MulCapRowMerge1.Columns(8).HeaderText = "负荷"
        DgV_MulCapRowMerge1.Columns(9).HeaderText = "备注"
        'DgV_MulCapRowMerge1.ColumnSCaption = strCap
        'DgV_MulCapRowMerge1.ColumnDeep = 1

        DgV_MulCapRowMerge1.ColumnHeadersHeight = 40
        'For i As Integer = 1 To DgV_MulCapRowMerge1.Columns.Count - 1
        '    DgV_MulCapRowMerge1.Columns(i).Width = 80
        'Next

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        SetButtonEnableStatus(False, False, False)
        DgV_MulCapRowMerge1.Columns.Clear()

        If RadioButton1.Checked Then
            ViewHourData()
        End If
        If RadioButton2.Checked Then
            ViewDailyData()
        End If

        If RadioButton3.Checked Then
            ViewMonthData()
        End If


        If RadioButton1.Checked Then
            OutPutHourDataToExcel()
        End If

        If RadioButton2.Checked Then
            OutPutDailyDataToExcel()
        End If

        If RadioButton3.Checked Then
            OutPutMonthDataToExcel()
        End If
        SetButtonEnableStatus(True, True, True)

    End Sub

    Private Sub OutPutHourDataToExcel()
        Dim d As String(,)
        ReDim d(DgV_MulCapRowMerge1.Rows.Count - 1, DgV_MulCapRowMerge1.Columns.Count - 2 - 1)

        Dim recordCnt = DgV_MulCapRowMerge1.Rows.Count
        If recordCnt < 1 Then
            MessageBox.Show("当前查询记录为0,无法导出!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        For i As Integer = 0 To DgV_MulCapRowMerge1.Rows.Count - 1
            d(i, 0) = DgV_MulCapRowMerge1.Item(1, i).Value
            d(i, 1) = DgV_MulCapRowMerge1.Item(2, i).Value
            d(i, 2) = DgV_MulCapRowMerge1.Item(3, i).Value
            d(i, 3) = DgV_MulCapRowMerge1.Item(4, i).Value
            d(i, 4) = DgV_MulCapRowMerge1.Item(5, i).Value
            d(i, 5) = DgV_MulCapRowMerge1.Item(6, i).Value
            d(i, 6) = DgV_MulCapRowMerge1.Item(7, i).Value
            d(i, 7) = DgV_MulCapRowMerge1.Item(8, i).Value
            d(i, 8) = DgV_MulCapRowMerge1.Item(9, i).Value
            d(i, 9) = DgV_MulCapRowMerge1.Item(10, i).Value
            d(i, 10) = DgV_MulCapRowMerge1.Item(11, i).Value
            d(i, 11) = DgV_MulCapRowMerge1.Item(12, i).Value
            d(i, 12) = DgV_MulCapRowMerge1.Item(13, i).Value
            d(i, 13) = DgV_MulCapRowMerge1.Item(14, i).Value
        Next

        Dim savefile As New SaveFileDialog()

        With savefile
            .Filter = "xls(文件)*.xls|*.xls"
        End With

        savefile.FileName = Now.ToString("yyyyMMdd")

        If savefile.ShowDialog() = DialogResult.OK Then
            Dim xlApp As New Excel.Application
            Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(Application.StartupPath & "\OutputXls\HourData.xls")
            Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Sheets(1)

            '快速写入EXCEL表中
            xlWorkSheet.Activate()
            xlWorkSheet.Range("B6").Resize(d.Length / 14, 14).Value = d
            'xlWorkSheet.Cells(6, 2） = 12

            '前面SaveFileDialog对话框已进行文件名是否存在的校验
            '关闭SaveAs()的警告
            xlApp.DisplayAlerts = False
            xlWorkBook.SaveAs(savefile.FileName)

            '保存完毕后回复SaveAs()的警告
            xlApp.DisplayAlerts = True
            xlApp.Quit()
        End If

    End Sub

    Private Sub OutPutDailyDataToExcel()


        Dim d As String(,)
        ReDim d(DgV_MulCapRowMerge1.Rows.Count - 1, DgV_MulCapRowMerge1.Columns.Count - 2 - 1)
        Dim recordCnt = DgV_MulCapRowMerge1.Rows.Count
        If recordCnt < 1 Then
            MessageBox.Show("当前查询记录为0,无法导出!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        For i As Integer = 0 To DgV_MulCapRowMerge1.Rows.Count - 1
            d(i, 0) = DgV_MulCapRowMerge1.Item(1, i).Value
            d(i, 1) = DgV_MulCapRowMerge1.Item(2, i).Value
            d(i, 2) = DgV_MulCapRowMerge1.Item(3, i).Value
            d(i, 3) = DgV_MulCapRowMerge1.Item(4, i).Value
            d(i, 4) = DgV_MulCapRowMerge1.Item(5, i).Value
            d(i, 5) = DgV_MulCapRowMerge1.Item(6, i).Value
            d(i, 6) = DgV_MulCapRowMerge1.Item(7, i).Value
            d(i, 7) = DgV_MulCapRowMerge1.Item(8, i).Value
            d(i, 8) = DgV_MulCapRowMerge1.Item(9, i).Value
            d(i, 9) = DgV_MulCapRowMerge1.Item(10, i).Value
            d(i, 10) = DgV_MulCapRowMerge1.Item(11, i).Value
            ' d(i, 11) = DgV_MulCapRowMerge1.Item(12, i).Value  '备注一列，目前不添加
        Next

        Dim savefile As New SaveFileDialog()

        With savefile
            .Filter = "xls(文件)*.xls|*.xls"
        End With

        savefile.FileName = Now.ToString("yyyyMM")

        If savefile.ShowDialog() = DialogResult.OK Then
            Dim xlApp As New Excel.Application
            Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(Application.StartupPath & "\OutputXls\DailyData.xls")
            Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Sheets(1)

            '快速写入EXCEL表中
            xlWorkSheet.Activate()
            xlWorkSheet.Range("B6").Resize(d.Length / 11, 11).Value = d  ' d.Length / 12 - 1, 12
            xlApp.DisplayAlerts = False
            xlWorkBook.SaveAs(savefile.FileName)

            xlApp.DisplayAlerts = True
            xlApp.Quit()
        End If
    End Sub

    Private Sub OutPutMonthDataToExcel()
        Dim d As String(,)
        ReDim d(DgV_MulCapRowMerge1.Rows.Count - 1, DgV_MulCapRowMerge1.Columns.Count - 2)

        Dim recordCnt = DgV_MulCapRowMerge1.Rows.Count
        If recordCnt < 1 Then
            MessageBox.Show("当前查询记录为0,无法导出!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        For i As Integer = 0 To DgV_MulCapRowMerge1.Rows.Count - 1
            d(i, 0) = DgV_MulCapRowMerge1.Item(1, i).Value
            d(i, 1) = DgV_MulCapRowMerge1.Item(2, i).Value
            d(i, 2) = DgV_MulCapRowMerge1.Item(3, i).Value
            d(i, 3) = DgV_MulCapRowMerge1.Item(4, i).Value
            d(i, 4) = DgV_MulCapRowMerge1.Item(5, i).Value
            d(i, 5) = DgV_MulCapRowMerge1.Item(6, i).Value
            d(i, 6) = DgV_MulCapRowMerge1.Item(7, i).Value
            d(i, 7) = DgV_MulCapRowMerge1.Item(8, i).Value
            d(i, 8) = DgV_MulCapRowMerge1.Item(9, i).Value
            'd(i, 9) = DgV_MulCapRowMerge1.Item(10, i).Value
        Next

        Dim savefile As New SaveFileDialog()

        With savefile
            .Filter = "xls(文件)*.xls|*.xls"
        End With

        savefile.FileName = Now.ToString("yyyy")

        If savefile.ShowDialog() = DialogResult.OK Then
            Dim xlApp As New Excel.Application
            Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Open(Application.StartupPath & "\OutputXls\MonthData.xls")
            Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Sheets(1)

            '快速写入EXCEL表中
            xlWorkSheet.Activate()
            xlWorkSheet.Range("B5").Resize(d.Length / 9, 9).Value = d
            xlApp.DisplayAlerts = False
            xlWorkBook.SaveAs(savefile.FileName)

            xlApp.DisplayAlerts = True
            xlApp.Quit()
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        Me.Button2.Enabled = False
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Me.Button2.Enabled = False
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Me.Button2.Enabled = False
    End Sub
End Class