Imports Ini.Net

Public Class FormKayParameter

    ' Private listname As String() = {"甲烷(标干)"， "总烃(标干)", "总烃(湿基)", "氧气(标干)", "温度", "压力", "流速", "湿度", "甲烷(湿基)", "甲烷(湿基)", "非甲烷总烃(湿基)", "氧气(湿基)"}

    Private listname As String() = {"甲烷(标干)", "总烃(标干)", "非甲烷总烃(标干)", "颗粒物(标干)", "氧气(标干)", "温度", "压力", "流速", "湿度", "甲烷(湿基)", "总烃(湿基)", "非甲烷总烃(湿基)", "颗粒物(湿基)", "氧气(湿基)"}

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub

    Private Sub InitializeDisplay()
        '关键参数填充
        TextBox1.Text = keyParameter.area
        TextBox2.Text = keyParameter.press
        TextBox3.Text = keyParameter.codffO2
        TextBox4.Text = keyParameter.speed
        TextBox5.Text = keyParameter.pitot
        TextBox6.Text = keyParameter.limitO2

        '量程上下限填充
        TextBox9.Text = rangeParameter.tchlow
        TextBox10.Text = rangeParameter.tchup
        TextBox11.Text = rangeParameter.o2low
        TextBox12.Text = rangeParameter.o2up

        TextBox13.Text = rangeParameter.temprangelow
        TextBox14.Text = rangeParameter.temprangeup
        TextBox15.Text = rangeParameter.pressrangelow
        TextBox16.Text = rangeParameter.pressrangeup
        TextBox17.Text = rangeParameter.speedrangelow
        TextBox18.Text = rangeParameter.speedrangeup
        TextBox19.Text = rangeParameter.humirangelow
        TextBox20.Text = rangeParameter.humirangeup
        '颗粒物量程下限,上限
        TextBox40.Text = rangeParameter.particulaterangelow
        TextBox41.Text = rangeParameter.particulaterangeup


        TextBox29.Text = alarmParameter.ch4alarmlow
        TextBox30.Text = alarmParameter.ch4alarmup
        TextBox31.Text = alarmParameter.nmhcalarmlow
        TextBox32.Text = alarmParameter.nmhcalarmup
        TextBox33.Text = alarmParameter.tchalarmlow
        TextBox34.Text = alarmParameter.tchalarmup

        TextBox350.Text = alarmParameter.particulatealarmlow
        TextBox360.Text = alarmParameter.particulatealarmup

        If keyParameter.corverFlag = 1 Then
            CheckBox1.Checked = True
        Else
            CheckBox1.Checked = False
        End If

        TextBox21.Text = purgeSetting.gasPurgeTime
        TextBox22.Text = purgeSetting.gasPurgeHold
        TextBox23.Text = purgeSetting.pitotTime
        TextBox24.Text = purgeSetting.pitotHold

        If purgeSetting.purgeFlag Then
            CheckBox2.Checked = True
        Else
            CheckBox2.Checked = False
        End If

        TextBox25.Text = calibrationPara.zeroTime
        TextBox26.Text = calibrationPara.zeroConc
        TextBox27.Text = calibrationPara.rangeTime
        TextBox28.Text = calibrationPara.rangeConc

        '校零或校标定时器的时长
        'UI显示ini文件中的配置
        TextBox39.Text = calibrationPara.calibrationTime
        TextBox42.Text = calibrationPara.rangeResetTime
        TextBox43.Text = calibrationPara.platformResetTime

    End Sub

    Private Sub FormKayParameter_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeDisplay()

        InitializeCheckBox()

        InitializeUserInfo()

        InitializeLogin()
    End Sub

    Private Sub InitializeLogin()
        Dim sqlHelper As SQLiteHelper

        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        Dim sql As String = "SELECT UserName as 用户名,userRight as 用户权限 FROM UserInfo "

        sqlHelper.SelectToDt(sql)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If

        DataGridView1.DataSource = sqlHelper.DatasDt
        DataGridView1.AllowUserToAddRows = False
    End Sub

    Private Sub InitializeCheckBox()


        Dim strbuf As String = ReadConfigIni("界面显示", "显示参数", "1,1,1,1,1,1,1,1,1,1,1,1,1,1")

        Dim tokens As String() = strbuf.Split(",")

        For i As Integer = 0 To listname.Length - 1
            If tokens(i) = 1 Then
                CheckedListBox1.Items.Add(listname(i), True)
            Else
                CheckedListBox1.Items.Add(listname(i), False)
            End If

        Next

    End Sub

    Private Sub SaveKayParameterIni()
        Dim InitializeName As New IniFile(IniPathName)

        keyParameter.area = TextBox1.Text
        keyParameter.press = TextBox2.Text
        keyParameter.codffO2 = TextBox3.Text
        keyParameter.speed = TextBox4.Text
        keyParameter.pitot = TextBox5.Text
        keyParameter.limitO2 = TextBox6.Text

        If CheckBox1.Checked Then
            keyParameter.corverFlag = 1
        Else
            keyParameter.corverFlag = 0
        End If

        InitializeName.WriteString("关键参数", "烟道面积", keyParameter.area)
        InitializeName.WriteString("关键参数", "空气系数", keyParameter.codffO2)
        InitializeName.WriteString("关键参数", "大气压", keyParameter.press)
        InitializeName.WriteString("关键参数", "皮托管系数", keyParameter.pitot)
        InitializeName.WriteString("关键参数", "速度场系数", keyParameter.speed)
        InitializeName.WriteString("关键参数", "折算氧限值", keyParameter.limitO2)
        InitializeName.WriteString("关键参数", "是否折算", keyParameter.corverFlag)

    End Sub

    Private Sub SavePurgeSetIni()
        Dim InitializeName As New IniFile(IniPathName)

        purgeSetting.gasPurgeTime = TextBox21.Text
        purgeSetting.gasPurgeHold = TextBox22.Text
        purgeSetting.pitotTime = TextBox23.Text
        purgeSetting.pitotHold = TextBox24.Text

        Dim strbuf As String

        If CheckBox2.Checked Then
            purgeSetting.purgeFlag = True
            strbuf = "1" & "," & purgeSetting.gasPurgeTime & "," & purgeSetting.gasPurgeHold & "," & purgeSetting.pitotTime & "," & purgeSetting.pitotHold
        Else
            purgeSetting.purgeFlag = False
            strbuf = "0" & "," & purgeSetting.gasPurgeTime & "," & purgeSetting.gasPurgeHold & "," & purgeSetting.pitotTime & "," & purgeSetting.pitotHold
        End If

        InitializeName.WriteString("反吹设置", "反吹时间设置", strbuf)

        '写PLC寄存器
        WritePurgeParameter()

    End Sub

    Private Sub WritePurgeParameter()
        Dim buff As UInt16()
        Dim b As UInt16()
        Dim flag(0) As UInt16
        ReDim buff(4)

        If CheckBox2.Checked Then
            flag(0) = 1
        Else
            flag(0) = 0
        End If

        '是否反吹
        Array.Copy(flag, 0, buff, 0, 1)

        b = IntegerToUshort(purgeSetting.gasPurgeTime)
        Array.Copy(b, 1, buff, 1, 1)

        b = IntegerToUshort(purgeSetting.gasPurgeHold)
        Array.Copy(b, 1, buff, 2, 1)

        b = IntegerToUshort(purgeSetting.pitotTime)
        Array.Copy(b, 1, buff, 3, 1)

        b = IntegerToUshort(purgeSetting.pitotHold)
        Array.Copy(b, 1, buff, 4, 1)

        plcdata.writeCom = True
        plcdata.writeAnaly = 1
        'PLCCommu提供的写“反吹设置参数”地址
        plcdata.writeAdd = 100
        plcdata.writeData = buff

    End Sub

    Private Sub SaveCalibrationSetIni()
        Dim InitializeName As New IniFile(IniPathName)

        calibrationPara.zeroTime = TextBox25.Text
        calibrationPara.zeroConc = TextBox26.Text
        calibrationPara.rangeTime = TextBox27.Text
        calibrationPara.rangeConc = TextBox28.Text
        '校零或校标定时器的时长
        '保存配置到ini文件
        calibrationPara.calibrationTime = TextBox39.Text
        calibrationPara.rangeResetTime = TextBox42.Text
        calibrationPara.platformResetTime = TextBox43.Text

        Dim strbuf As String

        strbuf = calibrationPara.zeroTime & "," & calibrationPara.zeroConc & "," & calibrationPara.rangeTime & "," & calibrationPara.rangeConc & "," _
            & calibrationPara.calibrationTime & "," & calibrationPara.rangeResetTime & "," & calibrationPara.platformResetTime


        InitializeName.WriteString("校准设置", "校准参数", strbuf)
    End Sub

    Private Sub SaveKeyParameterLog()
        Dim sqlHelper As SQLiteHelper
        Dim list1 As New ArrayList
        Dim str As String = Now.ToString("yyyy-MM-dd HH:mm:ss")

        Dim note As Integer = 1

        sqlHelper = New SQLiteHelper(ConfigFilePath)
        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        list1.Add("insert into 关键参数记录(samplingtime ,SamplingArea ,Press ,CoeffO2,SPEED ,Pitol ,LimitO2  ,ConvertFlag ,REMARK) 
                  values('" & str & "'," & keyParameter.area & "," & keyParameter.press & "," & keyParameter.codffO2 & "," & keyParameter.speed & "," & keyParameter.pitot & "," & keyParameter.limitO2 & "," & keyParameter.corverFlag & ",'" & note & "');")

        sqlHelper.IUD(list1)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If
    End Sub

    Private Sub SaveRangeParameterIni()
        Dim InitializeName As New IniFile(IniPathName)
        Dim strbuf As String

        rangeParameter.tchlow = TextBox9.Text
        rangeParameter.tchup = TextBox10.Text
        rangeParameter.o2low = TextBox11.Text
        rangeParameter.o2up = TextBox12.Text
        rangeParameter.temprangelow = TextBox13.Text
        rangeParameter.temprangeup = TextBox14.Text
        rangeParameter.pressrangelow = TextBox15.Text
        rangeParameter.pressrangeup = TextBox16.Text
        rangeParameter.speedrangelow = TextBox17.Text
        rangeParameter.speedrangeup = TextBox18.Text
        rangeParameter.humirangelow = TextBox19.Text
        rangeParameter.humirangeup = TextBox20.Text

        rangeParameter.ch4alarmlow = TextBox29.Text
        rangeParameter.ch4alarmup = TextBox30.Text
        rangeParameter.nmhcalarmlow = TextBox31.Text
        rangeParameter.nmhcalarmup = TextBox32.Text
        rangeParameter.tchalarmlow = TextBox33.Text
        rangeParameter.tchalarmup = TextBox34.Text

        '颗粒物量程下限,上限
        rangeParameter.particulaterangelow = TextBox40.Text
        rangeParameter.particulaterangeup = TextBox41.Text


        strbuf = rangeParameter.tchup & "," & rangeParameter.tchlow & "," & rangeParameter.tchalarmup & "," & rangeParameter.tchalarmlow
        InitializeName.WriteString("量程设置", "总烃限值", strbuf)
        strbuf = rangeParameter.ch4up & "," & rangeParameter.ch4low & "," & rangeParameter.ch4alarmup & "," & rangeParameter.ch4alarmlow
        InitializeName.WriteString("量程设置", "甲烷限值", strbuf)
        strbuf = rangeParameter.nmhcup & "," & rangeParameter.nmhclow & "," & rangeParameter.nmhcalarmup & "," & rangeParameter.nmhcalarmlow
        InitializeName.WriteString("量程设置", "非甲烷限值", strbuf)
        strbuf = rangeParameter.o2up & "," & rangeParameter.o2low & "," & rangeParameter.o2alarmup & "," & rangeParameter.o2alarmlow
        InitializeName.WriteString("量程设置", "氧气限值", strbuf)
        strbuf = rangeParameter.temprangeup & "," & rangeParameter.temprangelow & "," & rangeParameter.temprangeup & "," & rangeParameter.temprangelow
        InitializeName.WriteString("量程设置", "温度限值", strbuf)
        strbuf = rangeParameter.pressrangeup & "," & rangeParameter.pressrangelow & "," & rangeParameter.pressrangeup & "," & rangeParameter.pressrangelow
        InitializeName.WriteString("量程设置", "压力限值", strbuf)
        strbuf = rangeParameter.speedrangeup & "," & rangeParameter.speedrangelow & "," & rangeParameter.speedrangeup & "," & rangeParameter.speedrangelow
        InitializeName.WriteString("量程设置", "动压限值", strbuf)
        strbuf = rangeParameter.humirangeup & "," & rangeParameter.humirangelow & "," & rangeParameter.humirangeup & "," & rangeParameter.humirangelow
        InitializeName.WriteString("量程设置", "湿度限值", strbuf)

        strbuf = rangeParameter.particulaterangeup & "," & rangeParameter.particulaterangelow & "," & rangeParameter.particulatealarmup & "," & rangeParameter.particulatealarmlow
        InitializeName.WriteString("量程设置", "颗粒物限值", strbuf)

    End Sub

    Private Sub SaveRangeParameterLog()
        Dim sqlHelper As SQLiteHelper
        Dim list1 As New ArrayList
        Dim str As String = Now.ToString("yyyy-MM-dd HH:mm:ss")

        Dim note As Integer = 1

        sqlHelper = New SQLiteHelper(ConfigFilePath)
        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        list1.Add("insert into 量程记录(samplingtime,tchlow,tchup,o2low,o2up,templow,tempup,presslow,pressup,speedlow,speedup,humilow,humiup ) 
                  values('" & str & "'," & rangeParameter.tchlow & "," & rangeParameter.tchup & "," & rangeParameter.o2low & "," & rangeParameter.o2up & ",
                  " & rangeParameter.temprangelow & "," & rangeParameter.temprangeup & "," & rangeParameter.pressrangelow & "," & rangeParameter.pressrangeup & ",
                  " & rangeParameter.speedrangelow & "," & rangeParameter.speedrangeup & "," & rangeParameter.humirangelow & "," & rangeParameter.humirangeup & ");")

        sqlHelper.IUD(list1)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim InitializeName As New IniFile(IniPathName)

        Dim str As String = "1,1,1,1,1,1,1,1,1,1,1,1,1,1"

        For i As Integer = 0 To CheckedListBox1.Items.Count - 1
            If i = 0 Then
                If CheckedListBox1.GetItemChecked(i) Then
                    str = "1"
                Else
                    str = "0"
                End If
            Else
                If CheckedListBox1.GetItemChecked(i) Then
                    str = str & "," & "1"
                Else
                    str = str & "," & "0"
                End If
            End If
        Next

        InitializeName.WriteString("界面显示", "显示参数", str)
        maindisplay.Invoke(str)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        SaveKeyParameterLog()

        SaveKayParameterIni()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click


        SaveRangeParameterIni()

        SaveRangeParameterLog()
        '向PLCCommu的串口写入量程数据
        WriteRange()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        SavePurgeSetIni()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        SaveCalibrationSetIni()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        SaveUserInfo()
    End Sub

    Private Sub WriteRange()
        Dim buff As UInt16()
        ReDim buff(15)

        Dim b As UInt16()
        b = FloatToUshort(rangeParameter.temprangelow)
        Array.Copy(b, 0, buff, 2, 2)

        b = FloatToUshort(rangeParameter.temprangeup)
        Array.Copy(b, 0, buff, 0, 2)

        b = FloatToUshort(rangeParameter.pressrangelow)
        Array.Copy(b, 0, buff, 6, 2)

        b = FloatToUshort(rangeParameter.pressrangeup)
        Array.Copy(b, 0, buff, 4, 2)

        b = FloatToUshort(rangeParameter.speedrangelow)
        Array.Copy(b, 0, buff, 10, 2)

        b = FloatToUshort(rangeParameter.speedrangeup)
        Array.Copy(b, 0, buff, 8, 2)

        b = FloatToUshort(rangeParameter.humirangelow)
        Array.Copy(b, 0, buff, 14, 2)

        b = FloatToUshort(rangeParameter.humirangeup)
        Array.Copy(b, 0, buff, 12, 2)

        plcdata.writeCom = True
        plcdata.writeAnaly = 1
        plcdata.writeAdd = 50
        plcdata.writeData = buff
    End Sub

    Private Sub SaveUserInfo()
        Dim InitializeName As New IniFile(IniPathName)

        senddatatoserver.UserName = TextBox7.Text
        senddatatoserver.UserMN = TextBox8.Text

        InitializeName.WriteString("站点信息", "站点名称", senddatatoserver.UserName)
        InitializeName.WriteString("站点信息", "MN号", senddatatoserver.UserMN)
    End Sub

    Private Sub InitializeUserInfo()

        TextBox7.Text = senddatatoserver.UserName
        TextBox8.Text = senddatatoserver.UserMN

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim sqlHelper As SQLiteHelper

        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If
        '读取当前记录查询总数
        Dim sqlCount As String = "SELECT  count(*) FROM UserInfo where UserName='" & TextBox35.Text & "' "
        sqlHelper.SelectToDt(sqlCount)
        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            MessageBox.Show("当前数据库查询失败!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        Dim readSqlRecordCnt = sqlHelper.DatasDt.Rows(0).Item(0)
        If readSqlRecordCnt <> 0 Then
            MessageBox.Show("当前已存在用户，无法重复添加用户!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        If TextBox35.Text = "" OrElse TextBox37.Text = "" OrElse ComboBox1.Text = "" Then
            MessageBox.Show("当前待添加用户信息不全,请补齐后重试!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim sql As New ArrayList
        sql.Add("insert into UserInfo(UserName,userRight,userPW) values ('" & TextBox35.Text & "','" & ComboBox1.Text & "','" & TextBox37.Text & "')")

        sqlHelper.IUD(sql)

        '显示当前最新的用户表
        InitializeLogin()

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        Dim recordCnt = DataGridView1.RowCount
        If recordCnt = 0 Then
            MessageBox.Show("当前用户表用户为空,本次删除操作取消!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim strbuf As String = DataGridView1.CurrentRow.Cells.Item(0).Value
        Dim sql As New ArrayList

        Dim sqlHelper As SQLiteHelper
        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        sql.Add("delete from userinfo where username='" & strbuf & "'")
        sqlHelper.IUD(sql)

        '显示当前最新的用户表
        InitializeLogin()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.CurrentRow Is Nothing Then Exit Sub

        TextBox35.Text = DataGridView1.CurrentRow.Cells.Item(0).Value
        ComboBox1.Text = DataGridView1.CurrentRow.Cells.Item(1).Value
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim recordCnt = DataGridView1.RowCount
        If recordCnt = 0 Then
            MessageBox.Show("当前用户表用户为空,本次操作取消!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Dim strbuf As String = DataGridView1.CurrentRow.Cells.Item(0).Value
        Dim sql As New ArrayList

        Dim sqlHelper As SQLiteHelper
        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        sql.Add("update userinfo set userpw='" & TextBox37.Text & "' where username='" & strbuf & "'")
        sqlHelper.IUD(sql)
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        Dim buff As UShort()

        If TextBox36.Text = "" Then Exit Sub

        o2data.writeAnaly = &H1
        o2data.writeAdd = &H256

        buff = FloatToUshort(TextBox36.Text * 10000)
        o2data.writeData = buff
        o2data.writeCom = True
    End Sub

    Private Function FloatToUshort(fdata As Single) As UShort()
        Dim u(1) As UShort
        Dim b As Byte() = BitConverter.GetBytes(fdata)

        u(1) = BitConverter.ToUInt16(b, 0)
        u(0) = BitConverter.ToUInt16(b, 2)

        Return u
    End Function

    Private Function IntegerToUshort(data As Integer) As UShort()
        Dim u(1) As UShort
        Dim b As Byte() = BitConverter.GetBytes(data)

        u(1) = BitConverter.ToUInt16(b, 0)
        u(0) = BitConverter.ToUInt16(b, 2)

        Return u
    End Function

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim buff As UShort()

        If TextBox38.Text = "" Then Exit Sub

        o2data.writeAnaly = &H1
        o2data.writeAdd = &H246

        buff = FloatToUshort(TextBox38.Text * 10000)
        o2data.writeData = buff
        o2data.writeCom = True
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click

        Dim result As DialogResult
        result = MessageBox.Show("是否开始校标？", "提示", MessageBoxButtons.OKCancel)
        If result = DialogResult.Cancel Then Exit Sub

        Dim buff As UShort()
        ReDim buff(0)
        o2data.writeAnaly = &H1
        o2data.writeAdd = &H314

        buff(0) = 5

        o2data.writeData = buff
        o2data.writeCom = True
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

        Dim result As DialogResult

        result = MessageBox.Show("是否开始校零？", "提示", MessageBoxButtons.OKCancel)

        If result = DialogResult.Cancel Then Exit Sub

        Dim buff As UShort()
        ReDim buff(0)
        o2data.writeAnaly = &H1
        o2data.writeAdd = &H314

        buff(0) = 6
        o2data.writeData = buff
        o2data.writeCom = True
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Dim InitializeName As New IniFile(IniPathName)
        Dim strbuf As String

        alarmParameter.ch4alarmlow = TextBox29.Text
        alarmParameter.ch4alarmup = TextBox30.Text
        alarmParameter.nmhcalarmlow = TextBox31.Text
        alarmParameter.nmhcalarmup = TextBox32.Text
        alarmParameter.tchalarmlow = TextBox33.Text
        alarmParameter.tchalarmup = TextBox34.Text

        alarmParameter.particulatealarmlow = TextBox350.Text
        alarmParameter.particulatealarmup = TextBox360.Text


        strbuf = alarmParameter.ch4alarmlow & "," & alarmParameter.ch4alarmup & "," & alarmParameter.ch4rangelow & "," & alarmParameter.ch4rangeup
        InitializeName.WriteString("报警设置", "甲烷", strbuf)
        strbuf = alarmParameter.tchalarmlow & "," & alarmParameter.tchalarmup & "," & alarmParameter.tchrangelow & "," & alarmParameter.tchrangeup
        InitializeName.WriteString("报警设置", "总烃", strbuf)
        strbuf = alarmParameter.nmhcalarmlow & "," & alarmParameter.nmhcalarmup & "," & alarmParameter.nmhcrangelow & "," & alarmParameter.nmhcrangeup
        InitializeName.WriteString("报警设置", "非甲烷总烃", strbuf)


        strbuf = alarmParameter.particulatealarmlow & "," & alarmParameter.particulatealarmup & "," & alarmParameter.particulaterangelow & "," & alarmParameter.particulaterangeup
        InitializeName.WriteString("报警设置", "颗粒物", strbuf)

    End Sub
End Class