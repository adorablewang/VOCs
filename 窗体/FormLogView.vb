Imports System.IO
Public Class FormLogView
    Private Sub FormLogView_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList

        ComboBox1.Items.Add("通讯日志")
        ComboBox1.Items.Add("报警日志")
        ComboBox1.Items.Add("系统状态")
        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim directory As DirectoryInfo

        Dim files As FileInfo()

        ListView1.Items.Clear()

        If ComboBox1.SelectedIndex = 0 Then
            directory = New DirectoryInfo(Application.StartupPath & "\Log\communication")
            files = directory.GetFiles()
            For Each file As FileInfo In files
                ListView1.Items.Add(file.Name)
            Next
        End If

        If ComboBox1.SelectedIndex = 1 Then
            directory = New DirectoryInfo(Application.StartupPath & "\Log\alarm")
            files = directory.GetFiles()
            For Each file As FileInfo In files
                ListView1.Items.Add(file.Name)
            Next
        End If

        If ComboBox1.SelectedIndex = 2 Then
            directory = New DirectoryInfo(Application.StartupPath & "\Log\Operation")
            files = directory.GetFiles()
            For Each file As FileInfo In files
                ListView1.Items.Add(file.Name)
            Next
        End If
    End Sub


    Private Sub ListView1_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles ListView1.ItemSelectionChanged
        Dim filepath As String = ""

        If ComboBox1.SelectedIndex = 0 Then filepath = Application.StartupPath & "\Log\communication\" & e.Item.Text
        If ComboBox1.SelectedIndex = 1 Then filepath = Application.StartupPath & "\Log\alarm\" & e.Item.Text
        If ComboBox1.SelectedIndex = 2 Then filepath = Application.StartupPath & "\Log\Operation\" & e.Item.Text

        Dim sr As StreamReader = New StreamReader(filepath)
        Dim line As String = ""
        Do While sr.Peek > 0
            line += sr.ReadLine & vbCrLf
        Loop
        sr.Close()
        sr = Nothing

        RichTextBox1.Text = line
    End Sub
End Class