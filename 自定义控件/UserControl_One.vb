Imports System.Drawing.Drawing2D

Public Class UserControl_One

    Private _displayNumber As Integer
    Private _displayName As String

    Private _displayColor As Color

    Private g As Graphics
    Public Property DisplayColor() As Color
        Get
            Return _displayColor
        End Get
        Set(ByVal value As Color)
            _displayColor = value
            Panel1.BackColor = value
            Panel2.BackColor = value
            Label1.BackColor = value
        End Set
    End Property

    Sub New(ByVal number As Integer, ByVal name As String)

        ' 此调用是设计器所必需的。
        InitializeComponent()
        _displayNumber = number
        _displayName = name
        ' 在 InitializeComponent() 调用之后添加任何初始化。
        Label1.Text = _displayName
        Label13.Text = _displayName
        If number = 1 Then
            Panel1.Visible = True
            Panel2.Visible = False
        Else
            Panel1.Visible = False
            Panel2.Visible = True
        End If

    End Sub



    ''' <summary>
    ''' 显示名称和单位
    ''' </summary>
    ''' <param name="name">显示的名称</param>
    ''' <param name="unit">显示的单位</param>
    Public Sub SetTypeUnit(ByVal name As String, ByVal unit As String)
        If _displayNumber = 1 Then
            If name <> "" Then Label16.Text = name
            Label12.Text = unit
        End If

        If _displayNumber = 3 Then
            Dim tokens() As String
            If name <> "" Then
                tokens = name.Split(",")
                Label4.Text = tokens(0)
                Label5.Text = tokens(1)
                Label7.Text = tokens(2)
            End If

            tokens = unit.Split(",")
            Label2.Text = tokens(0)
            Label10.Text = tokens(1)
            Label11.Text = tokens(2)
        End If
    End Sub

    ''' <summary>
    ''' 显示浓度值
    ''' </summary>
    ''' <param name="value">测量浓度值</param>
    Public Sub SetValue(ByVal value As String)
        If _displayNumber = 1 Then
            Label14.Text = value
        End If

        If _displayNumber = 3 Then
            Dim tokens() As String
            tokens = value.Split(",")
            Label3.Text = tokens(0)
            Label8.Text = tokens(1)
            Label9.Text = tokens(2)
        End If
    End Sub

    Public Sub DisplayName(ByVal value As String)
        If _displayNumber = 1 Then
            Label16.Text = value
        End If

        If _displayNumber = 3 Then
            Dim tokens() As String
            tokens = value.Split(",")
            Label4.Text = tokens(0)
            Label5.Text = tokens(1)
            Label7.Text = tokens(2)
        End If
    End Sub

    Private Sub UserControl_One_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Height = 154
        Panel2.Top = 0
        Panel2.Left = 0
        Panel2.Width = Me.Width
        Panel2.Height = 154
        Panel1.Top = 0
        Panel1.Left = 0
        Panel1.Width = Me.Width
        Panel1.Height = 154

        Label1.Width = Me.Width
        Label1.Left = 0

        Dim path As GraphicsPath = GetRoundedRectPath(Me.ClientRectangle, 20)
        Me.Region = New Region(path)
    End Sub

    Private Sub UserControl_One_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        'Dim gradBrush As New LinearGradientBrush(Me.ClientRectangle, Color.DodgerBlue, Color.Blue, LinearGradientMode.BackwardDiagonal)
        'e.Graphics.FillRectangle(gradBrush, Me.ClientRectangle)

        Dim path As GraphicsPath = CreateRoundedRectPath(Label1.ClientRectangle, 20, True, True, True, True)
        'Label1.Region = New Region(path)
    End Sub


    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs)
        'Dim gradBrush As New LinearGradientBrush(Panel2.ClientRectangle, Color.LightSeaGreen, Color.Turquoise, LinearGradientMode.BackwardDiagonal)
        'e.Graphics.FillRectangle(gradBrush, Panel2.ClientRectangle)
    End Sub


    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs)
        'Dim gradBrush As New LinearGradientBrush(Panel1.ClientRectangle, Color.LightSeaGreen, Color.Turquoise, LinearGradientMode.BackwardDiagonal)
        'e.Graphics.FillRectangle(gradBrush, Panel1.ClientRectangle)
    End Sub

    Public Sub DrawLine()
        Dim mPen As New Pen(Color.Red)
        mPen.Width = 2

        'g = Me.PictureBox1.CreateGraphics

        g = Me.Panel2.CreateGraphics

        g.TranslateTransform(0, 10)
        g.ScaleTransform(1, -1)
        mPen.Color = Color.White

        g.DrawLine(mPen, 5, 5, 100, 5)
    End Sub

End Class
