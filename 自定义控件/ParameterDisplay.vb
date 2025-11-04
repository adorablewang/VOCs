Imports System.Drawing.Drawing2D

Public Class ParameterDisplay

    Public Event ChangeColor()
    Public Event ChangeSetValue()
    Private _color As Color
    ''' <summary>
    ''' 面板颜色
    ''' </summary>
    ''' <returns></returns>
    Public Property AnalyzerColor() As Color
        Get
            Return _color
        End Get
        Set(ByVal value As Color)
            _color = value
            UiAnalogMeter1.BodyColor = _color
        End Set
    End Property

    Private _name As String
    ''' <summary>
    ''' 面板名称
    ''' </summary>
    ''' <returns></returns>
    Public Property AnalyzerName() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
            Label1.Text = value
        End Set
    End Property

    Private _settemp As String
    Public Property AnalyzerSetValue() As String
        Get
            Return _settemp
        End Get
        Set(ByVal value As String)
            _settemp = value
            Label5.Text = value
        End Set
    End Property

    ''' <summary>
    ''' 设置实际显示温度
    ''' </summary>
    ''' <param name="value"></param>
    Public Sub SetAnalyzerValue(value As String)

        Label4.Text = value
        UiAnalogMeter1.Value = value
    End Sub

    Private Sub ParameterDisplay_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim path As GraphicsPath = GetRoundedRectPath(Me.ClientRectangle, 20)
        Me.Region = New Region(path)

        UiAnalogMeter1.BodyColor = _color
    End Sub

    Private Sub ParameterDisplay_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint

    End Sub

    Private Sub UiAnalogMeter1_ValueChanged(sender As Object, e As EventArgs) Handles UiAnalogMeter1.ValueChanged

    End Sub

    Private Sub UiAnalogMeter1_DoubleClick(sender As Object, e As EventArgs) Handles UiAnalogMeter1.DoubleClick
        RaiseEvent ChangeColor()
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub Label5_DoubleClick(sender As Object, e As EventArgs) Handles Label5.DoubleClick
        RaiseEvent ChangeSetValue()
    End Sub
End Class
