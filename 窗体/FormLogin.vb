Public Class FormLogin

    ' TODO: 插入代码，以使用提供的用户名和密码执行自定义的身份验证
    ' (请参阅 https://go.microsoft.com/fwlink/?LinkId=35339)。  
    ' 随后自定义主体可附加到当前线程的主体，如下所示: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' 其中 CustomPrincipal 是用于执行身份验证的 IPrincipal 实现。
    ' 随后，My.User 将返回 CustomPrincipal 对象中封装的标识信息
    ' 如用户名、显示名等

    Private dt As New DataTable
    Public Property IsLoggedIn As Boolean = False

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click

        Dim iIndex As Integer

        For iIndex = 0 To dt.Rows.Count - 1
            If dt.Rows(iIndex).Item(0) = ComboBox1.Text Then
                If dt.Rows(iIndex).Item(2) = PasswordTextBox.Text Then
                    '保存用户类型
                    If TextBox1.Text.Equals("管理员") Then
                        UserRole = EnumDefine.UserRoles.Administrator
                    ElseIf TextBox1.Text.Equals("巡视员") Then
                        UserRole = EnumDefine.UserRoles.Inspector
                    Else
                        UserRole = EnumDefine.UserRoles.Unknown
                    End If

                    '点击“确定”正常关闭
                    Me.DialogResult = DialogResult.OK
                    '登录验证成功
                    IsLoggedIn = True
                    Me.Close()
                Else
                    MessageBox.Show("密码输入错误，请重新输入", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    PasswordTextBox.Focus()
                    PasswordTextBox.SelectionStart = 0
                    PasswordTextBox.SelectionLength = PasswordTextBox.TextLength
                End If

                Exit For
            End If
        Next
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub FormLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeLogin()
    End Sub

    Private Sub InitializeLogin()
        Dim sqlHelper As SQLiteHelper
        Dim iIndex As Integer

        sqlHelper = New SQLiteHelper(ConfigFilePath)

        If Not sqlHelper.SQLStatus Then
            Exit Sub
        End If

        Dim sql As String = "SELECT UserName,userright,userpw FROM UserInfo"

        sqlHelper.SelectToDt(sql)

        If Not sqlHelper.SQLStatus Then
            sqlHelper.Close()
            Exit Sub
        End If

        dt = sqlHelper.DatasDt

        For iIndex = 0 To sqlHelper.DatasDt.Rows.Count - 1
            ComboBox1.Items.Add(sqlHelper.DatasDt.Rows(iIndex).Item(0))
        Next

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim strbuf As String
        Dim iIndex As Integer

        strbuf = ComboBox1.SelectedItem.ToString()

        For iIndex = 0 To dt.Rows.Count - 1
            If dt.Rows(iIndex).Item(0) = strbuf Then
                TextBox1.Text = dt.Rows(iIndex).Item（1）
                PasswordTextBox.Focus()
                PasswordTextBox.SelectionStart = 0
                PasswordTextBox.SelectionLength = PasswordTextBox.TextLength
            End If
        Next

    End Sub


End Class
