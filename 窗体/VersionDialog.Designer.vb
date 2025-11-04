<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class VersionDialog
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.UiLabel1 = New Sunny.UI.UILabel()
        Me.UiLabel2 = New Sunny.UI.UILabel()
        Me.UiLabel3 = New Sunny.UI.UILabel()
        Me.UiLabel4 = New Sunny.UI.UILabel()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(416, 204)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 27)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 21)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "确定"
        Me.OK_Button.Visible = False
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 21)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "取消"
        Me.Cancel_Button.Visible = False
        '
        'UiLabel1
        '
        Me.UiLabel1.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.UiLabel1.Location = New System.Drawing.Point(144, 12)
        Me.UiLabel1.Name = "UiLabel1"
        Me.UiLabel1.Size = New System.Drawing.Size(225, 23)
        Me.UiLabel1.TabIndex = 1
        Me.UiLabel1.Text = "固定污染源VOCs在线监测系统"
        Me.UiLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'UiLabel2
        '
        Me.UiLabel2.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.UiLabel2.Location = New System.Drawing.Point(144, 49)
        Me.UiLabel2.Name = "UiLabel2"
        Me.UiLabel2.Size = New System.Drawing.Size(53, 23)
        Me.UiLabel2.TabIndex = 2
        Me.UiLabel2.Text = "版本"
        Me.UiLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'UiLabel3
        '
        Me.UiLabel3.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.UiLabel3.Location = New System.Drawing.Point(144, 93)
        Me.UiLabel3.Name = "UiLabel3"
        Me.UiLabel3.Size = New System.Drawing.Size(436, 23)
        Me.UiLabel3.TabIndex = 3
        Me.UiLabel3.Text = "版权所有（@）2000-2025 湖北方圆科学仪器股份有限公司"
        Me.UiLabel3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'UiLabel4
        '
        Me.UiLabel4.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.UiLabel4.Location = New System.Drawing.Point(204, 48)
        Me.UiLabel4.Name = "UiLabel4"
        Me.UiLabel4.Size = New System.Drawing.Size(133, 23)
        Me.UiLabel4.TabIndex = 8
        Me.UiLabel4.Text = "1.0.24.1205"
        Me.UiLabel4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.污染源VOCs.My.Resources.Resources.icon
        Me.PictureBox1.Location = New System.Drawing.Point(2, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(131, 60)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 7
        Me.PictureBox1.TabStop = False
        '
        'VersionDialog
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(574, 242)
        Me.Controls.Add(Me.UiLabel4)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.UiLabel3)
        Me.Controls.Add(Me.UiLabel2)
        Me.Controls.Add(Me.UiLabel1)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "VersionDialog"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "关于.."
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents UiLabel1 As Sunny.UI.UILabel
    Friend WithEvents UiLabel2 As Sunny.UI.UILabel
    Friend WithEvents UiLabel3 As Sunny.UI.UILabel
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents UiLabel4 As Sunny.UI.UILabel
End Class
