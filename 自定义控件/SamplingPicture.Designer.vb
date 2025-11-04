<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SamplingPicture
    Inherits System.Windows.Forms.UserControl

    'UserControl 重写释放以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.UiButton1 = New Sunny.UI.UIButton()
        Me.SuspendLayout()
        '
        'UiButton1
        '
        Me.UiButton1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.UiButton1.FillColor = System.Drawing.SystemColors.Control
        Me.UiButton1.FillColor2 = System.Drawing.SystemColors.Control
        Me.UiButton1.FillHoverColor = System.Drawing.SystemColors.Control
        Me.UiButton1.FillPressColor = System.Drawing.SystemColors.Control
        Me.UiButton1.FillSelectedColor = System.Drawing.SystemColors.Control
        Me.UiButton1.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.UiButton1.ForeColor = System.Drawing.Color.Black
        Me.UiButton1.Location = New System.Drawing.Point(36, 50)
        Me.UiButton1.MinimumSize = New System.Drawing.Size(1, 1)
        Me.UiButton1.Name = "UiButton1"
        Me.UiButton1.RectColor = System.Drawing.Color.Black
        Me.UiButton1.Size = New System.Drawing.Size(69, 410)
        Me.UiButton1.Style = Sunny.UI.UIStyle.Custom
        Me.UiButton1.TabIndex = 0
        Me.UiButton1.Text = "烟道"
        Me.UiButton1.TipsFont = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        '
        'SamplingPicture
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.UiButton1)
        Me.Name = "SamplingPicture"
        Me.Size = New System.Drawing.Size(935, 514)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents UiButton1 As Sunny.UI.UIButton
End Class
